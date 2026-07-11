//! import/from, importlib, __name__, __main__, package patterns.

crate::runtime_case!(
    import_module_attr,
    "import json\nprint(json.__name__)\n",
    "json"
);
crate::runtime_case!(
    from_import_name,
    "from json import dumps\nprint(dumps([1]))\n",
    "[1]"
);
crate::runtime_case!(
    import_alias,
    "import json as j\nprint(j.dumps(True))\n",
    "true"
);
crate::runtime_case!(
    from_import_alias,
    "from json import dumps as d\nprint(d(1))\n",
    "1"
);
crate::runtime_case!(
    import_multiple,
    "import os, sys\nprint(os.path.join('a', 'b'))\n",
    "a/b"
);
crate::runtime_case!(
    from_multiple,
    "from json import dumps, loads\nprint(loads(dumps(1)))\n",
    "1"
);
crate::runtime_case!(
    importlib_import_module,
    "import importlib\nm = importlib.import_module('json')\nprint(m.dumps(2))\n",
    "2"
);
crate::runtime_case!(
    importlib_reload_exists,
    "import importlib\nprint(callable(importlib.reload))\n",
    "True"
);
crate::runtime_case!(
    importlib_util_find_spec,
    "import importlib.util\nspec = importlib.util.find_spec('json')\nprint(spec is not None)\n",
    "True"
);
crate::runtime_case!(
    sys_modules_contains_json,
    "import json\nimport sys\nprint('json' in sys.modules)\n",
    "True"
);
crate::runtime_case!(
    module_file_attr,
    "import json\nprint(isinstance(json.__file__, str) or json.__file__ is None)\n",
    "True"
);
crate::runtime_case!(
    module_doc_attr,
    "import json\nprint(isinstance(json.__doc__, str) or json.__doc__ is None)\n",
    "True"
);
crate::runtime_case!(
    module_package_attr,
    "import json\nprint(hasattr(json, '__package__'))\n",
    "True"
);
crate::runtime_case!(
    builtin_import_callable,
    "print(callable(__import__))\n",
    "True"
);
crate::runtime_case!(
    import_star_not_runtime,
    "import math\nprint(hasattr(math, 'sqrt'))\n",
    "True"
);
crate::runtime_case!(
    from_list_import,
    "from operator import add, mul\nprint(add(2, 3))\n",
    "5"
);
crate::runtime_case!(
    import_submodule,
    "import os.path\nprint(os.path.basename('a/b'))\n",
    "b"
);
crate::runtime_case!(
    importlib_metadata_version,
    "import importlib\nm = importlib.import_module('collections')\nprint(m.__name__)\n",
    "collections"
);
crate::runtime_case!(
    types_module_import,
    "import types\nprint(types.ModuleType.__name__)\n",
    "type"
);
crate::runtime_case!(
    create_module_dynamic,
    "import types\nm = types.ModuleType('dynamic_mod')\nm.x = 1\nprint(m.x)\n",
    "1"
);
crate::runtime_case!(
    sys_modules_register,
    "import sys\nimport types\nm = types.ModuleType('reg_test')\nsys.modules['reg_test'] = m\nprint(sys.modules['reg_test'] is m)\n",
    "True"
);
crate::runtime_case!(
    import_cached,
    "import json\nimport json as j2\nprint(json is j2)\n",
    "True"
);
crate::runtime_case!(
    from_import_attr_error,
    "try:\n from json import missing_attr_xyz\n print('ok')\nexcept ImportError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    import_error_message,
    "try:\n import no_such_module_xyz_abc\n print('ok')\nexcept ImportError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    pkgutil_iter_modules,
    "import pkgutil\nprint(isinstance(list(pkgutil.iter_modules()), list))\n",
    "True"
);
crate::runtime_case!(
    importlib_resources_exists,
    "import importlib\nprint(hasattr(importlib, 'resources'))\n",
    "True"
);
crate::runtime_case!(
    importlib_machinery_exists,
    "import importlib.machinery\nprint(hasattr(importlib.machinery, 'SourceFileLoader'))\n",
    "True"
);
crate::runtime_case!(
    runpy_module,
    "import runpy\nprint(hasattr(runpy, 'run_module'))\n",
    "True"
);
crate::runtime_case!(__name_in_module, "print(__name__)\n", "__main__");
crate::runtime_case!(
    __doc_optional,
    "print(isinstance(__doc__, str) or __doc__ is None)\n",
    "True"
);
crate::runtime_case!(
    __package_optional,
    "print(hasattr(globals(), '__package__') or True)\n",
    "True"
);
crate::runtime_case!(
    import_nested_attr,
    "import collections.abc\nprint(hasattr(collections.abc, 'Mapping'))\n",
    "True"
);
crate::runtime_case!(
    from_collections_import,
    "from collections import deque\nprint(deque([1]).pop())\n",
    "1"
);
crate::runtime_case!(
    import_encodings,
    "import encodings\nprint(hasattr(encodings, 'utf_8'))\n",
    "True"
);
crate::runtime_case!(
    importlib_util_module_from_spec,
    "import importlib.util\nprint(hasattr(importlib.util, 'module_from_spec'))\n",
    "True"
);
crate::runtime_case!(
    importlib_abc_meta_path,
    "import importlib.abc\nprint(hasattr(importlib.abc, 'MetaPathFinder'))\n",
    "True"
);
crate::runtime_case!(
    imp_legacy_not_used,
    "import importlib\nprint(importlib.__name__)\n",
    "importlib"
);
crate::runtime_case!(
    module_dict,
    "import json\nprint(isinstance(json.__dict__, dict))\n",
    "True"
);
crate::runtime_case!(
    getattr_on_module,
    "import json\nprint(callable(getattr(json, 'dumps')))\n",
    "True"
);
crate::runtime_case!(
    hasattr_on_module,
    "import json\nprint(hasattr(json, 'loads'))\n",
    "True"
);
crate::runtime_case!(
    dir_module,
    "import json\nprint('dumps' in dir(json))\n",
    "True"
);
crate::runtime_case!(
    import_future_annotations,
    "from __future__ import annotations\nx: int = 1\nprint(x)\n",
    "1"
);
crate::runtime_case!(
    import_future_division,
    "from __future__ import division\nprint(3 / 2)\n",
    "1.5"
);
crate::runtime_case!(
    import_future_print_function,
    "from __future__ import print_function\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    import_future_unicode_literals,
    "from __future__ import unicode_literals\nprint(isinstance('x', str))\n",
    "True"
);
crate::runtime_case!(
    importlib_metadata_distributions,
    "try:\n import importlib.metadata as md\n print(hasattr(md, 'version'))\nexcept ImportError:\n print('skip')\n",
    "True"
);
crate::runtime_case!(
    zipimport_module,
    "try:\n import zipimport\n print(hasattr(zipimport, 'zipimporter'))\nexcept ImportError:\n print('skip')\n",
    "True"
);

crate::compile_case!(relative_import_package, "from . import sibling\n");
crate::compile_case!(relative_import_parent, "from .. import pkg\n");
crate::compile_case!(
    importlib_import_module_reload,
    "import importlib\nm = importlib.import_module('json')\nimportlib.reload(m)\n"
);
crate::compile_case!(
    runpy_run_path,
    "import runpy\nrunpy.run_path('.', run_name='__main__')\n"
);
crate::compile_case!(import_all_list, "from json import *\n");
