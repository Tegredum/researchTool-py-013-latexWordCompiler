# author: DSH
"""Keep automated-test scratch files inside the repository and dispose of them.

The acceptance workflow runs on a source checkout and must not create files
outside the working directory. pytest's ``tmp_path``/``tmpdir`` fixtures use
the system temp directory by default, which on the acceptance machine cannot be
read without elevation. Each session is redirected to its own directory under
the git-ignored ``outputs/pytest-tmp``.

Files created by ``tests/test_mathtype.py`` are deliberately minimal or invalid
Word packages used to exercise failure paths; they are not conversion output.
A successful session deletes its own run directory; a failing session keeps it
when it contains diagnostics. An explicit ``--basetemp`` is left to pytest's
normal handling.
"""
import os
import shutil
import uuid
from pathlib import Path

import pytest

_BASE_TEMP = Path(__file__).resolve().parents[1] / "outputs" / "pytest-tmp"
_session_root: Path | None = None
_clean_on_success = False


def _redirect(factory: pytest.TempPathFactory, root: Path) -> pytest.TempPathFactory:
	# A resolved base temp skips the system temp lookup, the ``--basetemp`` wipe
	# and the numbered-directory rotation that pytest otherwise performs.
	factory._basetemp = root
	return factory


def pytest_configure(config: pytest.Config) -> None:
	global _session_root, _clean_on_success
	if config.option.basetemp:
		# The user explicitly chose a location; keep pytest's usual semantics.
		_clean_on_success = False
		return
	_BASE_TEMP.mkdir(parents=True, exist_ok=True)
	# A per-run directory keeps concurrent pytest sessions from sharing (or
	# deleting) each other's scratch files.
	_session_root = _BASE_TEMP / f"run-{os.getpid()}-{uuid.uuid4().hex[:8]}"
	_session_root.mkdir(parents=True, exist_ok=True)
	_clean_on_success = True
	# ``tmp_path`` and ``tmpdir`` read ``config._tmp_path_factory``; replacing it
	# here covers them regardless of which plugin provides ``tmp_path_factory``.
	config._tmp_path_factory = _redirect(
		pytest.TempPathFactory.from_config(config, _ispytest=True), _session_root
	)


def pytest_sessionfinish(session: pytest.Session, exitstatus: int) -> None:
	del session
	if not _clean_on_success or _session_root is None:
		return
	try:
		has_diagnostics = any(_session_root.iterdir())
	except OSError:
		has_diagnostics = False
	if exitstatus == pytest.ExitCode.OK or not has_diagnostics:
		shutil.rmtree(_session_root, ignore_errors=True)
		try:
			_BASE_TEMP.rmdir()
		except OSError:
			pass


@pytest.fixture(scope="session")
def tmp_path_factory(request: pytest.FixtureRequest) -> pytest.TempPathFactory:
	factory = pytest.TempPathFactory.from_config(request.config, _ispytest=True)
	if request.config.option.basetemp:
		return factory
	root = _session_root if _session_root is not None else _BASE_TEMP
	return _redirect(factory, root)
