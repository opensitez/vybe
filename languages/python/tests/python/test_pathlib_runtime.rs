//! pathlib PurePath and Path operations.

crate::runtime_case!(
    pathlib_purepath_name,
    "from pathlib import PurePath\nprint(PurePath('a/b/c.txt').name)\n",
    "c.txt"
);
crate::runtime_case!(
    pathlib_purepath_stem,
    "from pathlib import PurePath\nprint(PurePath('a/b/c.txt').stem)\n",
    "c"
);
crate::runtime_case!(
    pathlib_purepath_suffix,
    "from pathlib import PurePath\nprint(PurePath('a/b/c.txt').suffix)\n",
    ".txt"
);
crate::runtime_case!(
    pathlib_purepath_parent,
    "from pathlib import PurePath\nprint(str(PurePath('a/b/c').parent))\n",
    "a/b"
);
crate::runtime_case!(
    pathlib_purepath_join,
    "from pathlib import PurePath\nprint(str(PurePath('a') / 'b' / 'c'))\n",
    "a/b/c"
);
crate::runtime_case!(
    pathlib_purepath_parts,
    "from pathlib import PurePath\nprint(PurePath('a/b/c').parts)\n",
    "('a', 'b', 'c')"
);
crate::runtime_case!(
    pathlib_purepath_anchor,
    "from pathlib import PurePath\nprint(PurePath('/tmp/x').anchor)\n",
    "/"
);
crate::runtime_case!(
    pathlib_pureposixpath,
    "from pathlib import PurePosixPath\nprint(str(PurePosixPath('a/b')))\n",
    "a/b"
);
crate::runtime_case!(
    pathlib_purewindowspath,
    "from pathlib import PureWindowsPath\nprint(PureWindowsPath('a/b').parts[-1])\n",
    "b"
);
crate::runtime_case!(
    pathlib_path_cwd,
    "from pathlib import Path\nprint(isinstance(Path('.').resolve(), Path))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_exists,
    "from pathlib import Path\nprint(Path('.').exists())\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_is_dir,
    "from pathlib import Path\nprint(Path('.').is_dir())\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_is_file,
    "from pathlib import Path\nprint(Path('missing_xyz_path').is_file())\n",
    "False"
);
crate::runtime_case!(
    pathlib_path_with_suffix,
    "from pathlib import PurePath\nprint(PurePath('a.txt').with_suffix('.md').suffix)\n",
    ".md"
);
crate::runtime_case!(
    pathlib_path_with_stem,
    "from pathlib import PurePath\nprint(PurePath('a.txt').with_stem('b').name)\n",
    "b.txt"
);
crate::runtime_case!(
    pathlib_path_with_name,
    "from pathlib import PurePath\nprint(PurePath('a/b').with_name('c').name)\n",
    "c"
);
crate::runtime_case!(
    pathlib_path_match,
    "from pathlib import PurePath\nprint(PurePath('a/b/c.py').match('*.py'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_relative_to,
    "from pathlib import PurePath\nprint(PurePath('a/b/c').relative_to('a'))\n",
    "b/c"
);
crate::runtime_case!(
    pathlib_path_is_relative_to,
    "from pathlib import PurePath\nprint(PurePath('a/b/c').is_relative_to('a'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_joinpath,
    "from pathlib import PurePath\nprint(PurePath('a').joinpath('b', 'c').name)\n",
    "c"
);
crate::runtime_case!(
    pathlib_path_read_text_mock,
    "from pathlib import Path\nprint(hasattr(Path('.'), 'read_text'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_write_text_mock,
    "from pathlib import Path\nprint(hasattr(Path('.'), 'write_text'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_mkdir_exists,
    "from pathlib import Path\nprint(hasattr(Path('.'), 'mkdir'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_glob,
    "from pathlib import Path\nprint(hasattr(Path('.'), 'glob'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_iterdir,
    "from pathlib import Path\nprint(callable(Path('.').iterdir))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_home,
    "from pathlib import Path\nprint(isinstance(Path.home(), Path))\n",
    "True"
);
crate::runtime_case!(
    pathlib_slashes,
    "from pathlib import PurePath\nprint(PurePath('a') / 'b' == PurePath('a/b'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_truediv,
    "from pathlib import PurePath\nprint((PurePath('/tmp') / 'x').name)\n",
    "x"
);
crate::runtime_case!(
    pathlib_suffixes,
    "from pathlib import PurePath\nprint(PurePath('archive.tar.gz').suffixes)\n",
    "['.tar', '.gz']"
);
crate::runtime_case!(
    pathlib_drive,
    "from pathlib import PureWindowsPath\nprint(PureWindowsPath('C:/a').drive)\n",
    "C:"
);
crate::runtime_case!(
    pathlib_root,
    "from pathlib import PurePosixPath\nprint(PurePosixPath('/a/b').root)\n",
    "/"
);
crate::runtime_case!(
    pathlib_as_posix,
    "from pathlib import PurePath\nprint(PurePath('a/b').as_posix())\n",
    "a/b"
);
crate::runtime_case!(
    pathlib_as_uri,
    "from pathlib import PurePath\nprint(PurePath('/a/b').as_uri().startswith('file:'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_from_uri,
    "from pathlib import PurePath\nprint(PurePath.from_uri('file:///tmp').parts[-1])\n",
    "tmp"
);
crate::runtime_case!(
    pathlib_is_absolute,
    "from pathlib import PurePosixPath\nprint(PurePosixPath('/a').is_absolute())\n",
    "True"
);
crate::runtime_case!(
    pathlib_is_reserved,
    "from pathlib import PureWindowsPath\nprint(PureWindowsPath('CON').is_reserved())\n",
    "True"
);
crate::runtime_case!(
    pathlib_compare,
    "from pathlib import PurePath\nprint(PurePath('a') == PurePath('a'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_hash,
    "from pathlib import PurePath\nprint(hash(PurePath('a')) == hash(PurePath('a')))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_stat,
    "from pathlib import Path\nprint(hasattr(Path('.').stat(), 'st_size'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_touch,
    "from pathlib import Path\nprint(callable(Path('x').touch))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_unlink,
    "from pathlib import Path\nprint(callable(Path('x').unlink))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_rename,
    "from pathlib import Path\nprint(callable(Path('x').rename))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_resolve,
    "from pathlib import Path\nprint(isinstance(Path('.').resolve(), Path))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_expanduser,
    "from pathlib import Path\nprint(isinstance(Path('~').expanduser(), Path))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_read_bytes,
    "from pathlib import Path\nprint(hasattr(Path('.'), 'read_bytes'))\n",
    "True"
);
crate::runtime_case!(
    pathlib_path_open,
    "from pathlib import Path\nprint(callable(Path('.').open))\n",
    "True"
);

crate::compile_case!(
    pathlib_rglob,
    "from pathlib import Path\nlist(Path('.').rglob('*.rs'))\n"
);
crate::compile_case!(
    pathlib_walk,
    "from pathlib import Path\n[p for p in Path('.').iterdir()]\n"
);
crate::compile_case!(
    pathlib_hardlink,
    "from pathlib import Path\nhasattr(Path('.'), 'hardlink_to')\n"
);
crate::compile_case!(
    pathlib_symlink,
    "from pathlib import Path\nhasattr(Path('.'), 'symlink_to')\n"
);
crate::compile_case!(
    pathlib_chmod,
    "from pathlib import Path\nhasattr(Path('.'), 'chmod')\n"
);
