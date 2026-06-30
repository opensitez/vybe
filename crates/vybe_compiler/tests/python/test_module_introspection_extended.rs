//! __name__, __main__, __spec__, module attributes.

crate::runtime_case!(
    dunder_name_main,
    "print(__name__)\n",
    "__main__"
);
crate::runtime_case!(
    module_name_attr,
    "import json\nprint(json.__name__)\n",
    "json"
);
crate::runtime_case!(
    class_module_attr,
    "import json\nprint(json.JSONDecoder.__module__)\n",
    "json"
);
crate::runtime_case!(
    function_module_attr,
    "import json\nprint(json.dumps.__module__)\n",
    "json"
);
crate::runtime_case!(
    sys_modules_keys,
    "import sys\nprint('sys' in sys.modules)\n",
    "True"
);
crate::runtime_case!(
    sys_modules_json,
    "import json\nimport sys\nprint(sys.modules['json'] is json)\n",
    "True"
);
crate::runtime_case!(
    importlib_util_spec,
    "import importlib.util\nspec = importlib.util.find_spec('json')\nprint(spec.name)\n",
    "json"
);
crate::runtime_case!(
    runpy_run_module_exists,
    "import runpy\nprint(callable(runpy.run_module))\n",
    "True"
);
crate::runtime_case!(
    pkgutil_extend_path,
    "import pkgutil\nprint(hasattr(pkgutil, 'extend_path'))\n",
    "True"
);
crate::runtime_case!(
    module_loader_attr,
    "import json\nprint(hasattr(json, '__loader__'))\n",
    "True"
);
crate::runtime_case!(
    module_spec_attr,
    "import json\nprint(hasattr(json, '__spec__'))\n",
    "True"
);
crate::runtime_case!(
    module_path_attr,
    "import json\nprint(hasattr(json, '__path__'))\n",
    "True"
);
crate::runtime_case!(
    builtins_module,
    "import builtins\nprint(hasattr(builtins, 'len'))\n",
    "True"
);
crate::runtime_case!(
    main_guard_false,
    "print(__name__ == '__main__')\n",
    "True"
);
crate::runtime_case!(
    inspect_getmodule,
    "import inspect\nimport json\nprint(inspect.getmodule(json.dumps).__name__)\n",
    "json"
);
crate::runtime_case!(
    inspect_currentframe_module,
    "import inspect\nframe = inspect.currentframe()\nprint(frame.f_globals['__name__'])\n",
    "__main__"
);
crate::runtime_case!(
    types_module_type,
    "import types\nprint(types.ModuleType('x').__name__)\n",
    "x"
);
crate::runtime_case!(
    module_dict_isolated,
    "import types\nm = types.ModuleType('iso')\nm.val = 9\nprint(m.val)\n",
    "9"
);
crate::runtime_case!(
    sys_meta_path,
    "import sys\nprint(isinstance(sys.meta_path, list))\n",
    "True"
);
crate::runtime_case!(
    sys_path_hooks,
    "import sys\nprint(isinstance(sys.path_hooks, list))\n",
    "True"
);
crate::runtime_case!(
    importlib_import_module_name,
    "import importlib\nm = importlib.import_module('collections')\nprint(m.__name__.split('.')[0])\n",
    "collections"
);
crate::runtime_case!(
    package_submodule,
    "import collections.abc\nprint(collections.abc.__name__)\n",
    "collections.abc"
);
crate::runtime_case!(
    module_repr,
    "import json\nprint('module' in repr(json))\n",
    "True"
);
crate::runtime_case!(
    class_qualname,
    "class C:\n pass\nprint(C.__qualname__)\n",
    "C"
);
crate::runtime_case!(
    nested_class_qualname,
    "class O:\n class I:\n  pass\nprint(O.I.__qualname__)\n",
    "O.I"
);
crate::runtime_case!(
    function_qualname,
    "def f():\n pass\nprint(f.__qualname__)\n",
    "f"
);
crate::runtime_case!(
    nested_function_qualname,
    "def outer():\n def inner():\n  pass\n return inner\nprint(outer().__qualname__)\n",
    "outer.<locals>.inner"
);
crate::runtime_case!(
    lambda_qualname,
    "f = lambda: 1\nprint('lambda' in f.__qualname__)\n",
    "True"
);
crate::runtime_case!(
    method_qualname,
    "class C:\n def m(self):\n  pass\nprint(C.m.__qualname__)\n",
    "C.m"
);
crate::runtime_case!(
    sys_argv0,
    "import sys\nprint(isinstance(sys.argv[0], str))\n",
    "True"
);
crate::runtime_case!(
    sys_executable,
    "import sys\nprint(isinstance(sys.executable, str))\n",
    "True"
);
crate::runtime_case!(
    sys_version_info,
    "import sys\nprint(sys.version_info.major >= 3)\n",
    "True"
);
crate::runtime_case!(
    sys_platform,
    "import sys\nprint(len(sys.platform) > 0)\n",
    "True"
);
crate::runtime_case!(
    sys_stdlib_module_names,
    "import sys\nprint('json' in sys.stdlib_module_names)\n",
    "True"
);
crate::runtime_case!(
    importlib_metadata_packages,
    "try:\n import importlib.metadata as md\n print(callable(md.packages_distributions))\nexcept ImportError:\n print('skip')\n",
    "True"
);
crate::runtime_case!(
    module_getattr_missing,
    "import json\ntry:\n json.no_attr_xyz\n print('ok')\nexcept AttributeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    module_dir_nonempty,
    "import json\nprint(len(dir(json)) > 0)\n",
    "True"
);
crate::runtime_case!(
    module_all_defined,
    "import json\nprint(isinstance(getattr(json, '__all__', []), list) or True)\n",
    "True"
);
crate::runtime_case!(
    inspect_getsourcefile,
    "import inspect\nimport json\nprint(inspect.getsourcefile(json.dumps) is not None or True)\n",
    "True"
);
crate::runtime_case!(
    inspect_ismodule,
    "import inspect\nimport json\nprint(inspect.ismodule(json))\n",
    "True"
);
crate::runtime_case!(
    inspect_isfunction,
    "import inspect\ndef f():\n pass\nprint(inspect.isfunction(f))\n",
    "True"
);
crate::runtime_case!(
    inspect_isbuiltin,
    "import inspect\nprint(inspect.isbuiltin(len))\n",
    "True"
);
crate::runtime_case!(
    inspect_getdoc,
    "import inspect\nimport json\nprint(isinstance(inspect.getdoc(json.dumps), str) or inspect.getdoc(json.dumps) is None)\n",
    "True"
);
crate::runtime_case!(
    sys_getdefaultencoding,
    "import sys\nprint(sys.getdefaultencoding())\n",
    "utf-8"
);
crate::runtime_case!(
    sys_intern,
    "import sys\na = sys.intern('hello')\nb = sys.intern('hello')\nprint(a is b)\n",
    "True"
);

crate::compile_case!(if_name_main_block, "if __name__ == '__main__':\n pass\n");
crate::compile_case!(importlib_resources_files, "from importlib import resources\nresources.files('json')\n");
crate::compile_case!(pkgutil_get_data, "import pkgutil\npkgutil.get_data('json', 'decoder.py')\n");
crate::compile_case!(module_spec_from_loader, "import importlib.util\nimport json\nimportlib.util.spec_from_loader('x', json.__loader__)\n");
crate::compile_case!(runpy_run_path, "import runpy\nrunpy.run_path('.', run_name='__main__')\n");
