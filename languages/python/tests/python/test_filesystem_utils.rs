//! tempfile, shutil, glob, fnmatch, argparse, configparser.

crate::runtime_case!(
    tempfile_mkstemp,
    "import tempfile\nfd, path = tempfile.mkstemp()\nimport os\nos.close(fd)\nos.remove(path)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    tempfile_mkdtemp,
    "import tempfile\nimport shutil\npath = tempfile.mkdtemp()\nshutil.rmtree(path)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    tempfile_named_temporary,
    "import tempfile\nf = tempfile.NamedTemporaryFile(delete=False)\nf.write(b'hi')\nf.close()\nimport os\nos.remove(f.name)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    tempfile_gettempdir,
    "import tempfile\nprint(isinstance(tempfile.gettempdir(), str))\n",
    "True"
);
crate::runtime_case!(
    shutil_copyfileobj,
    "import shutil\nimport io\nsrc = io.BytesIO(b'data')\ndst = io.BytesIO()\nshutil.copyfileobj(src, dst)\nprint(dst.getvalue())\n",
    "b'data'"
);
crate::runtime_case!(
    shutil_copymode_exists,
    "import shutil\nprint(callable(shutil.copymode))\n",
    "True"
);
crate::runtime_case!(
    shutil_copystat_exists,
    "import shutil\nprint(callable(shutil.copystat))\n",
    "True"
);
crate::runtime_case!(
    shutil_disk_usage,
    "import shutil\nprint(shutil.disk_usage('.').free > 0)\n",
    "True"
);
crate::runtime_case!(
    shutil_which,
    "import shutil\nprint(shutil.which('python') is not None or shutil.which('python3') is not None)\n",
    "True"
);
crate::runtime_case!(
    glob_glob,
    "import glob\nprint(isinstance(glob.glob('*.rs'), list) or isinstance(glob.glob('*'), list))\n",
    "True"
);
crate::runtime_case!(
    glob_iglob,
    "import glob\nprint(hasattr(glob, 'iglob'))\n",
    "True"
);
crate::runtime_case!(
    glob_escape,
    "import glob\nprint(isinstance(glob.escape('*.py'), str))\n",
    "True"
);
crate::runtime_case!(
    fnmatch_fnmatch,
    "import fnmatch\nprint(fnmatch.fnmatch('test.py', '*.py'))\n",
    "True"
);
crate::runtime_case!(
    fnmatch_filter,
    "import fnmatch\nprint(fnmatch.filter(['a.py', 'b.txt'], '*.py'))\n",
    "['a.py']"
);
crate::runtime_case!(
    fnmatch_translate,
    "import fnmatch\nprint(fnmatch.translate('*.py'))\n",
    "(?s:.*\\.py)\\Z"
);
crate::runtime_case!(
    argparse_parser,
    "import argparse\np = argparse.ArgumentParser()\nprint(isinstance(p, argparse.ArgumentParser))\n",
    "True"
);
crate::runtime_case!(
    argparse_parse_args,
    "import argparse\np = argparse.ArgumentParser()\nargs = p.parse_args([])\nprint(args)\n",
    "Namespace()"
);
crate::runtime_case!(
    argparse_add_argument,
    "import argparse\np = argparse.ArgumentParser()\np.add_argument('--foo')\nargs = p.parse_args(['--foo', 'bar'])\nprint(args.foo)\n",
    "bar"
);
crate::runtime_case!(
    argparse_store_true,
    "import argparse\np = argparse.ArgumentParser()\np.add_argument('-v', action='store_true')\nargs = p.parse_args(['-v'])\nprint(args.v)\n",
    "True"
);
crate::runtime_case!(
    configparser_basic,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[sec]\\nkey = value\\n')\nprint(c['sec']['key'])\n",
    "value"
);
crate::runtime_case!(
    configparser_sections,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[a]\\nx=1\\n[b]\\ny=2\\n')\nprint(sorted(c.sections()))\n",
    "['a', 'b']"
);
crate::runtime_case!(
    configparser_has_option,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[s]\\nk=v\\n')\nprint(c.has_option('s', 'k'))\n",
    "True"
);
crate::runtime_case!(
    configparser_getint,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[s]\\nn = 42\\n')\nprint(c.getint('s', 'n'))\n",
    "42"
);
crate::runtime_case!(
    configparser_getboolean,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[s]\\nflag = true\\n')\nprint(c.getboolean('s', 'flag'))\n",
    "True"
);
crate::runtime_case!(
    tempfile_spooled_temporary,
    "import tempfile\nprint(hasattr(tempfile, 'SpooledTemporaryFile'))\n",
    "True"
);
crate::runtime_case!(
    shutil_rmtree_exists,
    "import shutil\nprint(callable(shutil.rmtree))\n",
    "True"
);
crate::runtime_case!(
    shutil_move_exists,
    "import shutil\nprint(callable(shutil.move))\n",
    "True"
);
crate::runtime_case!(
    shutil_copy_exists,
    "import shutil\nprint(callable(shutil.copy))\n",
    "True"
);
crate::runtime_case!(
    glob_has_magic,
    "import glob\nprint(glob.has_magic('*.py'))\n",
    "True"
);
crate::runtime_case!(
    fnmatch_casefold,
    "import fnmatch\nprint(fnmatch.fnmatchcase('A.py', '*.py'))\n",
    // `*` matches any sequence and involves no letter case, so real python3
    // says True here regardless of fnmatchcase being case-sensitive.
    "True"
);
crate::runtime_case!(
    argparse_namespace,
    "import argparse\nprint(hasattr(argparse, 'Namespace'))\n",
    "True"
);
crate::runtime_case!(
    configparser_defaults,
    "import configparser\nc = configparser.ConfigParser()\nprint(isinstance(c.defaults(), dict))\n",
    "True"
);
crate::runtime_case!(
    tempfile_tempdir,
    "import tempfile\nprint(tempfile.tempdir is None or isinstance(tempfile.tempdir, str))\n",
    "True"
);
crate::runtime_case!(
    shutil_get_terminal_size,
    "import shutil\nprint(shutil.get_terminal_size().columns > 0)\n",
    "True"
);
crate::runtime_case!(
    glob_brace,
    "import glob\nprint(isinstance(glob.glob('*'), list))\n",
    "True"
);
crate::runtime_case!(
    fnmatch_fnmatchcase,
    "import fnmatch\nprint(fnmatch.fnmatchcase('test.py', 'test.py'))\n",
    "True"
);
crate::runtime_case!(
    argparse_subparsers,
    "import argparse\np = argparse.ArgumentParser()\nsp = p.add_subparsers()\nprint(sp is not None)\n",
    "True"
);
crate::runtime_case!(
    configparser_interpolation,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[s]\\na = hello\\nb = ${a}\\n')\nprint(c['s']['b'])\n",
    "hello"
);
crate::runtime_case!(
    tempfile_mkstemp_suffix,
    "import tempfile\nfd, path = tempfile.mkstemp(suffix='.txt')\nimport os\nos.close(fd)\nprint(path.endswith('.txt'))\n",
    "True"
);
crate::runtime_case!(
    shutil_copy2_exists,
    "import shutil\nprint(callable(shutil.copy2))\n",
    "True"
);
crate::runtime_case!(
    glob_escape_no_magic,
    "import glob\nprint(glob.escape('file.txt'))\n",
    "file.txt"
);
crate::runtime_case!(
    fnmatch_translate_star,
    "import fnmatch\nprint('.*' in fnmatch.translate('*'))\n",
    "True"
);
crate::runtime_case!(
    argparse_mutually_exclusive,
    "import argparse\np = argparse.ArgumentParser()\ng = p.add_mutually_exclusive_group()\nprint(g is not None)\n",
    "True"
);
crate::runtime_case!(
    configparser_read_dict,
    "import configparser\nc = configparser.ConfigParser()\nc.read_dict({'s': {'k': 'v'}})\nprint(c['s']['k'])\n",
    "v"
);
crate::runtime_case!(
    tempfile_temporarydirectory,
    "import tempfile\nprint(hasattr(tempfile, 'TemporaryDirectory'))\n",
    "True"
);
crate::runtime_case!(
    shutil_chown_exists,
    "import shutil\nprint(hasattr(shutil, 'chown'))\n",
    "True"
);

crate::compile_case!(
    shutil_unpack_archive,
    "import shutil\nshutil.unpack_archive\n"
);
crate::compile_case!(
    tempfile_temporarydirectory_ctx,
    "import tempfile\nwith tempfile.TemporaryDirectory() as d:\n pass\n"
);
crate::compile_case!(
    argparse_filetype,
    "import argparse\nargparse.FileType('r')\n"
);
crate::compile_case!(
    configparser_raw,
    "import configparser\nc = configparser.ConfigParser()\nc.read_string('[s]\\nk=%(name)s\\n', source='s')\n"
);
crate::compile_case!(
    glob_recursive,
    "import glob\nglob.glob('**/*', recursive=True)\n"
);
