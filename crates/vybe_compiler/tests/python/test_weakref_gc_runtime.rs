//! weakref, gc, tracemalloc memory introspection.

crate::runtime_case!(
    weakref_ref_alive,
    "import weakref\nclass C: pass\no = C()\nr = weakref.ref(o)\nprint(r() is o)\n",
    "True"
);
crate::runtime_case!(
    weakref_ref_callable,
    "import weakref\nclass C: pass\nr = weakref.ref(C())\nprint(callable(r))\n",
    "True"
);
crate::runtime_case!(
    weakref_proxy_str,
    "import weakref\nclass C:\n def __str__(self):\n  return 'obj'\no = C()\np = weakref.proxy(o)\nprint(str(p))\n",
    "obj"
);
crate::runtime_case!(
    weakref_finalize,
    "import weakref\nclass C: pass\ncalled = []\nf = weakref.finalize(C(), lambda: called.append(1))\nprint(f.alive)\n",
    "True"
);
crate::runtime_case!(
    weakref_finalize_atexit,
    "import weakref\nclass C: pass\nf = weakref.finalize(C(), print)\nprint(f.atexit)\n",
    "True"
);
crate::runtime_case!(
    weakref_getweakrefcount,
    "import weakref\nclass C: pass\no = C()\nweakref.ref(o)\nprint(weakref.getweakrefcount(o) >= 1)\n",
    "True"
);
crate::runtime_case!(
    weakref_getweakrefs,
    "import weakref\nclass C: pass\no = C()\nweakref.ref(o)\nprint(len(weakref.getweakrefs(o)) >= 1)\n",
    "True"
);
crate::runtime_case!(
    weakref_proxy_getattr,
    "import weakref\nclass C:\n x = 5\np = weakref.proxy(C())\nprint(p.x)\n",
    "5"
);
crate::runtime_case!(
    weakref_ref_hash,
    "import weakref\nclass C: pass\nr = weakref.ref(C())\nprint(isinstance(hash(r), int))\n",
    "True"
);
crate::runtime_case!(
    weakref_ref_type,
    "import weakref\nclass C: pass\nr = weakref.ref(C())\nprint(type(r).__name__)\n",
    "ref"
);
crate::runtime_case!(
    weakref_proxy_type,
    "import weakref\nclass C: pass\np = weakref.proxy(C())\nprint(type(p).__name__)\n",
    "ProxyType"
);
crate::runtime_case!(
    weakref_callback,
    "import weakref\nlog = []\nclass C: pass\ndef cb(r):\n log.append(1)\no = C()\nweakref.ref(o, cb)\nprint(len(log))\n",
    "0"
);
crate::runtime_case!(
    gc_collect,
    "import gc\nprint(isinstance(gc.collect(), int))\n",
    "True"
);
crate::runtime_case!(
    gc_enable_disable,
    "import gc\ngc.disable()\ngc.enable()\nprint(gc.isenabled())\n",
    "True"
);
crate::runtime_case!(
    gc_isenabled,
    "import gc\nprint(isinstance(gc.isenabled(), bool))\n",
    "True"
);
crate::runtime_case!(
    gc_get_count,
    "import gc\nprint(len(gc.get_count()) == 3)\n",
    "True"
);
crate::runtime_case!(
    gc_get_threshold,
    "import gc\nprint(len(gc.get_threshold()) == 3)\n",
    "True"
);
crate::runtime_case!(
    gc_set_threshold,
    "import gc\nold = gc.get_threshold()\ngc.set_threshold(*old)\nprint(gc.get_threshold() == old)\n",
    "True"
);
crate::runtime_case!(
    gc_get_objects,
    "import gc\nprint(len(gc.get_objects()) > 0)\n",
    "True"
);
crate::runtime_case!(
    gc_get_referrers,
    "import gc\nx = []\nprint(isinstance(gc.get_referrers(x), list))\n",
    "True"
);
crate::runtime_case!(
    gc_get_referents,
    "import gc\nx = [1, 2]\nprint(isinstance(gc.get_referents(x), list))\n",
    "True"
);
crate::runtime_case!(
    gc_is_tracked,
    "import gc\nprint(gc.is_tracked([]))\n",
    "True"
);
crate::runtime_case!(
    gc_is_tracked_int,
    "import gc\nprint(gc.is_tracked(1))\n",
    "False"
);
crate::runtime_case!(
    gc_garbage_list,
    "import gc\nprint(isinstance(gc.garbage, list))\n",
    "True"
);
crate::runtime_case!(
    gc_get_stats,
    "import gc\nprint(isinstance(gc.get_stats(), list))\n",
    "True"
);
crate::runtime_case!(
    tracemalloc_is_tracing,
    "import tracemalloc\nprint(tracemalloc.is_tracing())\n",
    "False"
);
crate::runtime_case!(
    tracemalloc_start_stop,
    "import tracemalloc\ntracemalloc.start()\nprint(tracemalloc.is_tracing())\n",
    "True"
);
crate::runtime_case!(
    tracemalloc_get_traced_memory,
    "import tracemalloc\ntry:\n tracemalloc.start()\n print(tracemalloc.get_traced_memory()[0] >= 0)\nfinally:\n tracemalloc.stop()\n",
    "True"
);
crate::runtime_case!(
    tracemalloc_get_object_traceback,
    "import tracemalloc\nprint(hasattr(tracemalloc, 'get_object_traceback'))\n",
    "True"
);
crate::runtime_case!(
    weakref_weakset,
    "import weakref\nclass C: pass\ns = weakref.WeakSet()\ns.add(C())\nprint(len(s))\n",
    "1"
);
crate::runtime_case!(
    weakref_weakkeydictionary,
    "import weakref\nclass C: pass\nd = weakref.WeakKeyDictionary()\nc = C()\nd[c] = 1\nprint(d[c])\n",
    "1"
);
crate::runtime_case!(
    weakref_weakvaluedictionary,
    "import weakref\nclass C: pass\nd = weakref.WeakValueDictionary()\nc = C()\nd['k'] = c\nprint(d['k'] is c)\n",
    "True"
);
crate::runtime_case!(
    gc_set_debug,
    "import gc\ngc.set_debug(0)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    gc_get_debug,
    "import gc\nprint(isinstance(gc.get_debug(), int))\n",
    "True"
);
crate::runtime_case!(
    weakref_ref_eq,
    "import weakref\nclass C: pass\no = C()\nr1 = weakref.ref(o)\nr2 = weakref.ref(o)\nprint(r1 == r2)\n",
    "False"
);
crate::runtime_case!(
    weakref_proxy_eq,
    "import weakref\nclass C:\n def __eq__(self, o):\n  return True\np = weakref.proxy(C())\nprint(p == C())\n",
    "True"
);
crate::runtime_case!(
    gc_freeze_unfreeze,
    "import gc\nif hasattr(gc, 'freeze'):\n gc.freeze()\n gc.unfreeze()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    tracemalloc_snapshot,
    "import tracemalloc\ntry:\n tracemalloc.start()\n snap = tracemalloc.take_snapshot()\n print(len(snap.statistics('lineno')) >= 0)\nfinally:\n tracemalloc.stop()\n",
    "True"
);
crate::runtime_case!(
    weakref_module_name,
    "import weakref\nprint(weakref.__name__)\n",
    "weakref"
);
crate::runtime_case!(
    gc_module_name,
    "import gc\nprint(gc.__name__)\n",
    "gc"
);
crate::runtime_case!(
    tracemalloc_module_name,
    "import tracemalloc\nprint(tracemalloc.__name__)\n",
    "tracemalloc"
);
crate::runtime_case!(
    weakref_proxy_del,
    "import weakref\nclass C: pass\np = weakref.proxy(C())\nprint(callable(p.__class__))\n",
    "True"
);
crate::runtime_case!(
    gc_collect_returns_int,
    "import gc\nn = gc.collect()\nprint(n >= 0)\n",
    "True"
);
crate::runtime_case!(
    weakref_finalize_detach,
    "import weakref\nclass C: pass\nf = weakref.finalize(C(), print)\nf.detach()\nprint(f.alive)\n",
    "False"
);

crate::compile_case!(weakref_proxy_del_attr, "import weakref\nclass C:\n x = 1\np = weakref.proxy(C())\ndel p.x\n");
crate::compile_case!(gc_get_callbacks, "import gc\ngc.callbacks\n");
crate::compile_case!(tracemalloc_compare_snapshots, "import tracemalloc\ntracemalloc.start()\ns1 = tracemalloc.take_snapshot()\ns2 = tracemalloc.take_snapshot()\ntracemalloc.stop()\n");
crate::compile_case!(weakref_callable_proxy, "import weakref\nclass C:\n def __call__(self):\n  return 1\np = weakref.proxy(C())\np()\n");
crate::compile_case!(gc_get_objects_filter, "import gc\n[o for o in gc.get_objects() if isinstance(o, list)][:1]\n");
