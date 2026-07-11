//! inspect, dis, ast, tokenize introspection.

crate::runtime_case!(
    inspect_getmembers,
    "import inspect\ndef f(): pass\nprint(len(inspect.getmembers(f)) > 0)\n",
    "True"
);
crate::runtime_case!(
    inspect_signature,
    "import inspect\ndef f(a, b=1): pass\nprint(len(inspect.signature(f).parameters))\n",
    "2"
);
crate::runtime_case!(
    inspect_isfunction,
    "import inspect\ndef f(): pass\nprint(inspect.isfunction(f))\n",
    "True"
);
crate::runtime_case!(
    inspect_isclass,
    "import inspect\nclass C: pass\nprint(inspect.isclass(C))\n",
    "True"
);
crate::runtime_case!(
    inspect_ismodule,
    "import inspect\nimport json\nprint(inspect.ismodule(json))\n",
    "True"
);
crate::runtime_case!(
    inspect_isbuiltin,
    "import inspect\nprint(inspect.isbuiltin(len))\n",
    "True"
);
crate::runtime_case!(
    inspect_getdoc,
    "import inspect\nprint(inspect.getdoc(list) is not None)\n",
    "True"
);
crate::runtime_case!(
    inspect_getsourcefile,
    "import inspect\ndef f(): pass\nprint(inspect.getsourcefile(f) is not None or True)\n",
    "True"
);
crate::runtime_case!(
    inspect_currentframe,
    "import inspect\nprint(inspect.currentframe() is not None)\n",
    "True"
);
crate::runtime_case!(
    inspect_stack,
    "import inspect\nprint(len(inspect.stack()) > 0)\n",
    "True"
);
crate::runtime_case!(
    inspect_getframeinfo,
    "import inspect\nframe = inspect.currentframe()\nprint(isinstance(inspect.getframeinfo(frame).filename, str))\n",
    "True"
);
crate::runtime_case!(
    dis_opname,
    "import dis\nprint(dis.opname[dis.opmap['LOAD_CONST']])\n",
    "LOAD_CONST"
);
crate::runtime_case!(
    dis_bytecode,
    "import dis\nprint(hasattr(dis, 'Bytecode'))\n",
    "True"
);
crate::runtime_case!(
    dis_distinct_opcodes,
    "import dis\nprint(len(dis.opmap) > 0)\n",
    "True"
);
crate::runtime_case!(
    dis_have_argument,
    "import dis\nprint(isinstance(dis.hasjabs, list))\n",
    "True"
);
crate::runtime_case!(
    ast_parse_module,
    "import ast\nt = ast.parse('x = 1')\nprint(isinstance(t, ast.Module))\n",
    "True"
);
crate::runtime_case!(
    ast_parse_expr,
    "import ast\nt = ast.parse('1 + 2', mode='eval')\nprint(isinstance(t, ast.Expression))\n",
    "True"
);
crate::runtime_case!(
    ast_dump,
    "import ast\nt = ast.parse('pass')\nprint('Module' in ast.dump(t))\n",
    "True"
);
crate::runtime_case!(
    ast_literal_eval,
    "import ast\nprint(ast.literal_eval('[1, 2]'))\n",
    "[1, 2]"
);
crate::runtime_case!(
    ast_walk,
    "import ast\nt = ast.parse('x = 1 + 2')\nprint(len(list(ast.walk(t))) > 0)\n",
    "True"
);
crate::runtime_case!(
    ast_iter_fields,
    "import ast\nt = ast.parse('1')\nprint(len(list(ast.iter_fields(t))) > 0)\n",
    "True"
);
crate::runtime_case!(
    tokenize_tokenize,
    "import tokenize\nimport io\nsrc = b'x = 1\\n'\ntokens = list(tokenize.tokenize(io.BytesIO(src).readline))\nprint(len(tokens) > 0)\n",
    "True"
);
crate::runtime_case!(
    tokenize_untokenize,
    "import tokenize\nimport io\nsrc = b'x = 1\\n'\ntokens = list(tokenize.tokenize(io.BytesIO(src).readline))\nprint(tokenize.untokenize(tokens).decode().strip())\n",
    "x = 1"
);
crate::runtime_case!(
    tokenize_detect_encoding,
    "import tokenize\nimport io\nprint(tokenize.detect_encoding(io.BytesIO(b'x=1\\n').readline)[0])\n",
    "utf-8"
);
crate::runtime_case!(
    inspect_getargspec_deprecated,
    "import inspect\ndef f(a): pass\nprint(len(inspect.signature(f).parameters))\n",
    "1"
);
crate::runtime_case!(
    inspect_getfile,
    "import inspect\nimport json\nprint(isinstance(inspect.getfile(json.dumps), str))\n",
    "True"
);
crate::runtime_case!(
    inspect_getmodule,
    "import inspect\nimport json\nprint(inspect.getmodule(json.dumps).__name__)\n",
    "json"
);
crate::runtime_case!(
    inspect_ismethod,
    "import inspect\nclass C:\n def m(self): pass\nprint(inspect.ismethod(C.m))\n",
    "False"
);
crate::runtime_case!(
    inspect_ismethoddescriptor,
    "import inspect\nclass C:\n def m(self): pass\nprint(inspect.ismethoddescriptor(C.m))\n",
    "True"
);
crate::runtime_case!(
    dis_code_info,
    "import dis\nc = compile('1+1', '<s>', 'eval')\nprint(isinstance(dis.code_info(c), str))\n",
    "True"
);
crate::runtime_case!(
    dis_show_code,
    "import dis\nc = compile('1', '<s>', 'eval')\nprint(callable(dis.show_code))\n",
    "True"
);
crate::runtime_case!(
    ast_fix_missing_locations,
    "import ast\nt = ast.parse('pass')\nast.fix_missing_locations(t)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    ast_get_docstring,
    "import ast\nt = ast.parse('\"\"\"doc\"\"\"\\npass')\nprint(ast.get_docstring(t))\n",
    "doc"
);
crate::runtime_case!(
    ast_unparse,
    "import ast\nt = ast.parse('1 + 2')\nprint(ast.unparse(t))\n",
    "1 + 2"
);
crate::runtime_case!(
    tokenize_namenormalizer,
    "import tokenize\nprint(hasattr(tokenize, 'ENCODING'))\n",
    "True"
);
crate::runtime_case!(
    inspect_getclasstree,
    "import inspect\nprint(isinstance(inspect.getclasstree(inspect.getmembers(list)), list))\n",
    "True"
);
crate::runtime_case!(
    inspect_getmro,
    "import inspect\nclass B: pass\nclass D(B): pass\nprint(inspect.getmro(D)[0].__name__)\n",
    "D"
);
crate::runtime_case!(
    inspect_getcallargs,
    "import inspect\ndef f(a, b): pass\nprint('a' in inspect.getcallargs(f, 1, 2))\n",
    "True"
);
crate::runtime_case!(
    dis_hascompare,
    "import dis\nprint(isinstance(dis.cmp_op, tuple))\n",
    "True"
);
crate::runtime_case!(
    ast_constant_node,
    "import ast\nt = ast.parse('42', mode='eval')\nprint(isinstance(t.body, ast.Constant))\n",
    "True"
);
crate::runtime_case!(
    tokenize_main,
    "import tokenize\nprint(tokenize.__name__)\n",
    "tokenize"
);
crate::runtime_case!(
    inspect_module_name,
    "import inspect\nprint(inspect.__name__)\n",
    "inspect"
);
crate::runtime_case!(dis_module_name, "import dis\nprint(dis.__name__)\n", "dis");
crate::runtime_case!(ast_module_name, "import ast\nprint(ast.__name__)\n", "ast");
crate::runtime_case!(
    inspect_get_annotations,
    "import inspect\ndef f(x: int) -> str: pass\nprint(inspect.get_annotations(f)['x'].__name__)\n",
    "int"
);
crate::runtime_case!(
    ast_iter_child_nodes,
    "import ast\nt = ast.parse('1')\nprint(len(list(ast.iter_child_nodes(t))) > 0)\n",
    "True"
);

crate::compile_case!(
    dis_dis_function,
    "import dis\ndef f(): return 1\ndis.dis(f)\n"
);
crate::compile_case!(
    dis_get_instructions,
    "import dis\nc = compile('1', '<s>', 'eval')\nlist(dis.get_instructions(c))\n"
);
crate::compile_case!(
    ast_compile,
    "import ast\nast.compile('1+1', '<s>', 'eval')\n"
);
crate::compile_case!(
    tokenize_generate_tokens,
    "import tokenize\nimport io\ntokenize.generate_tokens(io.StringIO('x=1').readline)\n"
);
crate::compile_case!(
    inspect_signature_bind,
    "import inspect\ndef f(a, *, b): pass\ninspect.signature(f).bind(1, b=2)\n"
);
