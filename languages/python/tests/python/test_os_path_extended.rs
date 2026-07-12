//! os.path extended: normpath, split, join variants, exists, isfile, isdir.

crate::runtime_case!(
    os_path_join_varargs,
    "import os\nprint(os.path.join('a', 'b', 'c'))\n",
    "a/b/c"
);
crate::runtime_case!(
    os_path_join_absolute_tail,
    "import os\nprint(os.path.join('/tmp', 'file.txt'))\n",
    "/tmp/file.txt"
);
crate::runtime_case!(
    os_path_split_pair,
    "import os\nprint(os.path.split('/tmp/data/file.txt'))\n",
    "('/tmp/data', 'file.txt')"
);
crate::runtime_case!(
    os_path_splitroot,
    "import os\nprint(hasattr(os.path, 'splitroot'))\n",
    "True"
);
crate::runtime_case!(
    os_path_splitext,
    "import os\nprint(os.path.splitext('archive.tar.gz'))\n",
    "('archive.tar', '.gz')"
);
crate::runtime_case!(
    os_path_basename,
    "import os\nprint(os.path.basename('/a/b/c.txt'))\n",
    "c.txt"
);
crate::runtime_case!(
    os_path_dirname,
    "import os\nprint(os.path.dirname('/a/b/c.txt'))\n",
    "/a/b"
);
crate::runtime_case!(
    os_path_normpath_dots,
    "import os\nprint(os.path.normpath('a/../b'))\n",
    "b"
);
crate::runtime_case!(
    os_path_abspath_relative,
    "import os\nprint(isinstance(os.path.abspath('.'), str))\n",
    "True"
);
crate::runtime_case!(
    os_path_isabs_unix,
    "import os\nprint(os.path.isabs('/tmp'))\n",
    "True"
);
crate::runtime_case!(
    os_path_isabs_relative,
    "import os\nprint(os.path.isabs('tmp'))\n",
    "False"
);
crate::runtime_case!(
    os_path_commonpath,
    "import os\nprint(os.path.commonpath(['/tmp/a', '/tmp/b']))\n",
    "/tmp"
);
crate::runtime_case!(
    os_path_relpath,
    "import os\nprint(os.path.relpath('/tmp/a/b', '/tmp/a'))\n",
    "b"
);
crate::runtime_case!(
    os_path_samefile_exists,
    "import os\nprint(hasattr(os.path, 'samefile'))\n",
    "True"
);
crate::runtime_case!(
    os_path_getsize_exists,
    "import os\nprint(hasattr(os.path, 'getsize'))\n",
    "True"
);
crate::runtime_case!(
    os_path_getmtime_exists,
    "import os\nprint(hasattr(os.path, 'getmtime'))\n",
    "True"
);
crate::runtime_case!(
    os_path_isfile_false,
    "import os\nprint(os.path.isfile('definitely_missing_file_xyz'))\n",
    "False"
);
crate::runtime_case!(
    os_path_isdir_false,
    "import os\nprint(os.path.isdir('definitely_missing_dir_xyz'))\n",
    "False"
);
crate::runtime_case!(
    os_path_exists_false,
    "import os\nprint(os.path.exists('missing_path_xyz'))\n",
    "False"
);
crate::runtime_case!(
    os_path_islink_false,
    "import os\nprint(os.path.islink('missing_path_xyz'))\n",
    "False"
);
crate::runtime_case!(
    os_path_ismount_root,
    "import os\nprint(os.path.ismount('/'))\n",
    "True"
);
crate::runtime_case!(
    os_path_expanduser_tilde,
    "import os\nprint(isinstance(os.path.expanduser('~'), str))\n",
    "True"
);
crate::runtime_case!(
    os_path_expandvars,
    "import os\nprint(isinstance(os.path.expandvars('$HOME'), str))\n",
    "True"
);
crate::runtime_case!(os_path_curdir, "import os\nprint(os.path.curdir)\n", ".");
crate::runtime_case!(os_path_pardir, "import os\nprint(os.path.pardir)\n", "..");
crate::runtime_case!(
    os_path_sep,
    "import os\nprint(isinstance(os.path.sep, str))\n",
    "True"
);
crate::runtime_case!(
    os_path_altsep,
    "import os\nprint(os.path.altsep is None or isinstance(os.path.altsep, str))\n",
    "True"
);
crate::runtime_case!(os_path_extsep, "import os\nprint(os.path.extsep)\n", ".");
crate::runtime_case!(
    os_path_pathsep,
    "import os\nprint(isinstance(os.path.pathsep, str))\n",
    "True"
);
crate::runtime_case!(
    os_path_defpath,
    "import os\nprint(isinstance(os.path.defpath, str))\n",
    "True"
);
crate::runtime_case!(
    os_path_devnull,
    "import os\nprint(isinstance(os.path.devnull, str))\n",
    "True"
);
crate::runtime_case!(
    os_path_lexists,
    "import os\nprint(hasattr(os.path, 'lexists'))\n",
    "True"
);
crate::runtime_case!(
    os_path_realpath,
    "import os\nprint(isinstance(os.path.realpath('.'), str))\n",
    "True"
);
crate::runtime_case!(
    os_path_splitdrive,
    "import os\nprint(hasattr(os.path, 'splitdrive'))\n",
    "True"
);
crate::runtime_case!(
    os_path_join_empty,
    "import os\nprint(os.path.join(''))\n",
    ""
);
crate::runtime_case!(
    os_path_normcase,
    "import os\nprint(isinstance(os.path.normcase('ABC'), str))\n",
    "True"
);
crate::runtime_case!(
    os_path_commonprefix,
    "import os\nprint(os.path.commonprefix(['/tmp/a', '/tmp/b']))\n",
    "/tmp/"
);
crate::runtime_case!(
    os_path_basename_no_sep,
    "import os\nprint(os.path.basename('file.txt'))\n",
    "file.txt"
);
crate::runtime_case!(
    os_path_dirname_no_sep,
    "import os\nprint(os.path.dirname('file.txt'))\n",
    ""
);
crate::runtime_case!(
    os_path_splitext_no_ext,
    "import os\nprint(os.path.splitext('README'))\n",
    "('README', '')"
);
crate::runtime_case!(
    os_path_isabs_empty,
    "import os\nprint(os.path.isabs(''))\n",
    "False"
);
crate::runtime_case!(
    os_path_join_with_dot,
    "import os\nprint(os.path.join('a', '.', 'b'))\n",
    "a/./b"
);
crate::runtime_case!(
    os_path_relpath_same,
    "import os\nprint(os.path.relpath('a', 'a'))\n",
    "."
);
crate::runtime_case!(
    os_path_getatime_exists,
    "import os\nprint(hasattr(os.path, 'getatime'))\n",
    "True"
);
crate::runtime_case!(
    os_path_getctime_exists,
    "import os\nprint(hasattr(os.path, 'getctime'))\n",
    "True"
);
crate::runtime_case!(
    os_path_sameopenfile_exists,
    "import os\nprint(hasattr(os.path, 'sameopenfile'))\n",
    "True"
);
crate::runtime_case!(
    os_path_isblock_exists,
    "import os\nprint(hasattr(os.path, 'isblock'))\n",
    "True"
);
crate::runtime_case!(
    os_path_ischar_exists,
    "import os\nprint(hasattr(os.path, 'ischar'))\n",
    "True"
);
crate::runtime_case!(
    os_path_isfifo_exists,
    "import os\nprint(hasattr(os.path, 'isfifo'))\n",
    "True"
);
crate::runtime_case!(
    os_path_issocket_exists,
    "import os\nprint(hasattr(os.path, 'issocket'))\n",
    "True"
);

crate::compile_case!(
    os_path_walk,
    "import os\nfor root, dirs, files in os.walk('.'):\n break\n"
);
crate::compile_case!(os_path_scandir, "import os\nlist(os.scandir('.'))\n");
crate::compile_case!(
    os_path_mkdir,
    "import os\nos.mkdir('tmp_test_dir', 0o755)\n"
);
crate::compile_case!(os_path_symlink, "import os\nhasattr(os, 'symlink')\n");
crate::compile_case!(os_path_readlink, "import os\nhasattr(os, 'readlink')\n");
