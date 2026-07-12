//! Coroutine extended tests — creation, wrapping, status, yielding, errors (Lua 5.x §2.6)

lua_print! {
    co_create_status => {
        "local co = coroutine.create(function() end)\nprint(coroutine.status(co))\n",
        "suspended"
    },
    co_running_main => {
        "local running, is_main = coroutine.running()\nprint(type(running), is_main)\n",
        "thread\ttrue"
    },
    co_resume_arguments => {
        "local co = coroutine.create(function(a, b) print(a .. b) end)\ncoroutine.resume(co, \"x\", \"y\")\n",
        "xy"
    },
    co_yield_values => {
        "local co = coroutine.create(function() coroutine.yield(\"a\", \"b\") end)\nlocal _, r1, r2 = coroutine.resume(co)\nprint(r1, r2)\n",
        "a\tb"
    },
    co_resume_yield_result => {
        "local co = coroutine.create(function() local x = coroutine.yield() print(x) end)\ncoroutine.resume(co)  -- run to yield\ncoroutine.resume(co, 99)\n",
        "99"
    },
    co_wrap_callable => {
        "local f = coroutine.wrap(function(x) return x + 1 end)\nprint(f(10))\n",
        "11"
    },
    co_error_recovery => {
        "local co = coroutine.create(function() error(\"crash\") end)\nlocal ok, err = coroutine.resume(co)\nprint(ok, type(err))\n",
        "false\tstring"
    },
    co_status_dead => {
        "local co = coroutine.create(function() end)\ncoroutine.resume(co)\nprint(coroutine.status(co))\n",
        "dead"
    },
    co_status_normal_call => {
        "local co1, co2\nco1 = coroutine.create(function()\n  coroutine.resume(co2)\nend)\nco2 = coroutine.create(function()\n  print(coroutine.status(co1))\nend)\ncoroutine.resume(co1)\n",
        "normal"
    },
    co_yield_main_error => {
        "local ok = pcall(coroutine.yield)\nprint(ok)\n",
        "false"
    },
}
