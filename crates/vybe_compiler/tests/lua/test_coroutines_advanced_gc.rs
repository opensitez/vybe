//! Coroutines and garbage collection interactions (Lua 5.x §2.5, §2.6)

lua_print! {
    coroutine_gc_status_remains_dead => {
        "local co = coroutine.create(function() end)\ncoroutine.resume(co)\ncollectgarbage()\nprint(coroutine.status(co))\n",
        "dead"
    },
    coroutine_gc_weak_keys => {
        "local t = setmetatable({}, {__mode=\"k\"})\nlocal co = coroutine.create(function() end)\nt[co] = 42\nprint(t[co])\n",
        "42"
    },
    coroutine_gc_weak_values => {
        "local t = setmetatable({}, {__mode=\"v\"})\nlocal co = coroutine.create(function() end)\nt[1] = co\nprint(type(t[1]))\n",
        "thread"
    },
}
