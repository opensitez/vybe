//! pickle and marshal serialization roundtrips.

crate::runtime_case!(
    pickle_dumps_loads_list,
    "import pickle\ndata = [1, 2, 3]\nprint(pickle.loads(pickle.dumps(data)))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    pickle_dumps_loads_dict,
    "import pickle\ndata = {'a': 1, 'b': 2}\nprint(pickle.loads(pickle.dumps(data)))\n",
    "{'a': 1, 'b': 2}"
);
crate::runtime_case!(
    pickle_dumps_loads_tuple,
    "import pickle\nprint(pickle.loads(pickle.dumps((1, 2))))\n",
    "(1, 2)"
);
crate::runtime_case!(
    pickle_dumps_loads_str,
    "import pickle\nprint(pickle.loads(pickle.dumps('hello')))\n",
    "hello"
);
crate::runtime_case!(
    pickle_dumps_loads_int,
    "import pickle\nprint(pickle.loads(pickle.dumps(42)))\n",
    "42"
);
crate::runtime_case!(
    pickle_dumps_loads_none,
    "import pickle\nprint(pickle.loads(pickle.dumps(None)))\n",
    "None"
);
crate::runtime_case!(
    pickle_dumps_loads_bool,
    "import pickle\nprint(pickle.loads(pickle.dumps(True)))\n",
    "True"
);
crate::runtime_case!(
    pickle_dumps_loads_set,
    "import pickle\ndata = {1, 2, 3}\nprint(sorted(pickle.loads(pickle.dumps(data))))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    pickle_nested,
    "import pickle\ndata = {'items': [1, {'k': 2}]}\nprint(pickle.loads(pickle.dumps(data))['items'][1]['k'])\n",
    "2"
);
crate::runtime_case!(
    pickle_protocol_default,
    "import pickle\nprint(isinstance(pickle.dumps([1]), bytes))\n",
    "True"
);
crate::runtime_case!(
    pickle_protocol_highest,
    "import pickle\nprint(len(pickle.dumps([1], protocol=pickle.HIGHEST_PROTOCOL)) > 0)\n",
    "True"
);
crate::runtime_case!(
    pickle_load_dump_roundtrip,
    "import pickle\nimport io\nbuf = io.BytesIO()\npickle.dump([1, 2], buf)\nbuf.seek(0)\nprint(pickle.load(buf))\n",
    "[1, 2]"
);
crate::runtime_case!(
    marshal_dumps_loads_int,
    "import marshal\nprint(marshal.loads(marshal.dumps(99)))\n",
    "99"
);
crate::runtime_case!(
    marshal_dumps_loads_str,
    "import marshal\nprint(marshal.loads(marshal.dumps('hi')))\n",
    "hi"
);
crate::runtime_case!(
    marshal_dumps_loads_list,
    "import marshal\nprint(marshal.loads(marshal.dumps([1, 2])))\n",
    "[1, 2]"
);
crate::runtime_case!(
    marshal_dumps_loads_dict,
    "import marshal\nprint(marshal.loads(marshal.dumps({'x': 1})))\n",
    "{'x': 1}"
);
crate::runtime_case!(
    marshal_dumps_loads_tuple,
    "import marshal\nprint(marshal.loads(marshal.dumps((1,))))\n",
    "(1,)"
);
crate::runtime_case!(
    marshal_version,
    "import marshal\nprint(isinstance(marshal.version, int))\n",
    "True"
);
crate::runtime_case!(
    pickle_bytes_type,
    "import pickle\nb = pickle.dumps(1)\nprint(type(b).__name__)\n",
    "bytes"
);
crate::runtime_case!(
    pickle_float,
    "import pickle\nprint(pickle.loads(pickle.dumps(3.14)))\n",
    "3.14"
);
crate::runtime_case!(
    pickle_empty_list,
    "import pickle\nprint(pickle.loads(pickle.dumps([])))\n",
    "[]"
);
crate::runtime_case!(
    pickle_empty_dict,
    "import pickle\nprint(pickle.loads(pickle.dumps({})))\n",
    "{}"
);
crate::runtime_case!(
    pickle_bytes_obj,
    "import pickle\nprint(pickle.loads(pickle.dumps(b'abc')))\n",
    "b'abc'"
);
crate::runtime_case!(
    pickle_list_nested,
    "import pickle\nprint(pickle.loads(pickle.dumps([[1], [2]])))\n",
    "[[1], [2]]"
);
crate::runtime_case!(
    pickle_dict_keys_preserved,
    "import pickle\nd = {'a': 1, 'b': 2}\nprint(sorted(pickle.loads(pickle.dumps(d)).keys()))\n",
    "['a', 'b']"
);
crate::runtime_case!(
    marshal_float,
    "import marshal\nprint(marshal.loads(marshal.dumps(1.5)))\n",
    "1.5"
);
crate::runtime_case!(
    marshal_bool,
    "import marshal\nprint(marshal.loads(marshal.dumps(False)))\n",
    "False"
);
crate::runtime_case!(
    marshal_none,
    "import marshal\nprint(marshal.loads(marshal.dumps(None)))\n",
    "None"
);
crate::runtime_case!(
    pickle_copyreg_exists,
    "import copyreg\nprint(callable(copyreg.pickle))\n",
    "True"
);
crate::runtime_case!(
    pickle_pickletools_exists,
    "import pickletools\nprint(hasattr(pickletools, 'dis'))\n",
    "True"
);
crate::runtime_case!(
    pickle_unpickleable_error,
    "import pickle\nclass C:\n pass\ntry:\n pickle.dumps(C())\n print('ok')\nexcept (pickle.PicklingError, TypeError):\n print('err')\n",
    "err"
);
crate::runtime_case!(
    pickle_loads_bytes_only,
    "import pickle\ntry:\n pickle.loads('not bytes')\n print('ok')\nexcept (TypeError, pickle.UnpicklingError):\n print('err')\n",
    "err"
);
crate::runtime_case!(
    marshal_nested,
    "import marshal\nprint(marshal.loads(marshal.dumps({'a': [1, 2]})))\n",
    "{'a': [1, 2]}"
);
crate::runtime_case!(
    pickle_idempotent_json_like,
    "import pickle\ns = pickle.dumps({'k': [1]})\nprint(pickle.dumps(pickle.loads(s)) is not None)\n",
    "True"
);
crate::runtime_case!(
    marshal_large_int,
    "import marshal\nprint(marshal.loads(marshal.dumps(10 ** 18)))\n",
    "1000000000000000000"
);
crate::runtime_case!(
    pickle_unicode,
    "import pickle\nprint(pickle.loads(pickle.dumps('é')))\n",
    "é"
);
crate::runtime_case!(
    marshal_unicode,
    "import marshal\nprint(marshal.loads(marshal.dumps('é')))\n",
    "é"
);
crate::runtime_case!(
    pickle_module_name,
    "import pickle\nprint(pickle.__name__)\n",
    "pickle"
);
crate::runtime_case!(
    marshal_module_name,
    "import marshal\nprint(marshal.__name__)\n",
    "marshal"
);
crate::runtime_case!(
    pickle_default_protocol,
    "import pickle\nprint(isinstance(pickle.DEFAULT_PROTOCOL, int))\n",
    "True"
);
crate::runtime_case!(
    pickle_highest_protocol,
    "import pickle\nprint(pickle.HIGHEST_PROTOCOL >= pickle.DEFAULT_PROTOCOL)\n",
    "True"
);
crate::runtime_case!(
    pickle_loads_empty_bytes_raises,
    "import pickle\ntry:\n pickle.loads(b'')\n print('ok')\nexcept pickle.UnpicklingError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    marshal_empty_bytes_raises,
    "import marshal\ntry:\n marshal.loads(b'')\n print('ok')\nexcept (ValueError, EOFError):\n print('err')\n",
    "err"
);
crate::runtime_case!(
    pickle_roundtrip_frozenset,
    "import pickle\nprint(pickle.loads(pickle.dumps(frozenset({1, 2}))))\n",
    "frozenset({1, 2})"
);
crate::runtime_case!(
    pickle_roundtrip_range_not_picklable,
    "import pickle\ntry:\n pickle.dumps(range(3))\n print('ok')\nexcept (TypeError, pickle.PicklingError):\n print('err')\n",
    "err"
);

crate::compile_case!(
    pickle_pickleable_objects,
    "import pickle\nclass C:\n pass\ntry:\n pickle.dumps(C())\nexcept:\n pass\n"
);
crate::compile_case!(
    marshal_code_object,
    "import marshal\ncompile('1+1', '<s>', 'eval')\n"
);
crate::compile_case!(
    pickle_multiprocessing_reducer,
    "import pickle\npickle.Pickler\n"
);
crate::compile_case!(
    pickle_dbm_persist,
    "import pickle\nimport io\npickle.Pickler(io.BytesIO())\n"
);
crate::compile_case!(
    marshal_read_write_file,
    "import marshal\nimport tempfile\nf = tempfile.NamedTemporaryFile()\nmarshal.dump(1, f)\n"
);
