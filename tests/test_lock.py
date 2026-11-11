"""Single-user lock on the shared folder, simulating two PCs."""

import json
import os

import pytest


@pytest.fixture
def two_pcs(tool, monkeypatch):
    monkeypatch.setattr(tool.time, "sleep", lambda s: None)

    def as_pc(host, pid):
        monkeypatch.setattr(tool, "_HOST", host)
        monkeypatch.setattr(tool.os, "getpid", lambda: pid)
    return as_pc


def test_second_pc_is_refused_while_first_is_alive(tool, two_pcs):
    two_pcs("pcA", 1); assert tool.acquire_lock()
    two_pcs("pcB", 2); assert not tool.acquire_lock()
    two_pcs("pcA", 1); assert tool.refresh_lock()


def test_takeover_is_noticed(tool, two_pcs):
    two_pcs("pcA", 1); tool.acquire_lock()
    two_pcs("pcB", 2); tool.release_lock(force=True); assert tool.acquire_lock()
    two_pcs("pcA", 1); assert not tool.refresh_lock()
    tool.release_lock()                       # must not remove B's lock
    assert json.loads(tool.lock_file().read_text())["host"] == "pcB"


def test_empty_lock_file_counts_as_held(tool, two_pcs):
    two_pcs("pcA", 1); tool.acquire_lock()
    tool.lock_file().write_text("")           # half-synced by OneDrive
    two_pcs("pcB", 2); assert not tool.acquire_lock()
    old = tool.time.time() - tool.LOCK_STALE_SECONDS - 60
    os.utime(tool.lock_file(), (old, old))
    assert tool.acquire_lock()


def test_stale_lock_from_another_pc_is_taken(tool, two_pcs):
    two_pcs("pcA", 1); tool.acquire_lock()
    info = json.loads(tool.lock_file().read_text())
    info["heartbeat_epoch"] -= tool.LOCK_STALE_SECONDS + 60
    tool.lock_file().write_text(json.dumps(info))
    two_pcs("pcB", 2); assert tool.acquire_lock()


def test_own_lock_unreadable_while_onedrive_syncs_it(tool, two_pcs, monkeypatch):
    """OneDrive holds the lock just written: the user must not be locked out by their own lock."""
    two_pcs("pcA", 1)
    real = type(tool.lock_file()).read_text
    calls = {"n": 0}

    def busy(self, *a, **k):
        if self.name == tool.lock_file().name:
            calls["n"] += 1
            raise PermissionError("in use by OneDrive")
        return real(self, *a, **k)
    monkeypatch.setattr(type(tool.lock_file()), "read_text", busy)
    assert tool.acquire_lock()                # no lock, write, read back fails: ours
    assert calls["n"] >= 5                    # it tried several times first


def test_lock_message_says_why(tool, two_pcs):
    two_pcs("pcA", 1); tool.acquire_lock()
    assert "pcA" in tool.read_lock_info()
    tool.lock_file().write_text("")
    assert "unreadable lock file" in tool.read_lock_info() and "OneDrive" in tool.read_lock_info()
    tool.lock_file().unlink()
    assert "nobody" in tool.read_lock_info()
