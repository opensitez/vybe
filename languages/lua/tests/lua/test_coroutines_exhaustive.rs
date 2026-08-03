//! Coroutine library exhaustive tests: create, resume, yield, status, wrap, nested yielding (Lua 5.x §2.6)

lua_print! {
    co_exh_create_status => {
        "local co = coroutine.create(function() end)\nprint(coroutine.status(co))\n",
        "suspended"
    },
    co_exh_running_main => {
        "local running, is_main = coroutine.running()\nprint(type(running), is_main)\n",
        "thread\ttrue"
    },
    co_exh_resume_args => {
        "local co = coroutine.create(function(a, b) print(a .. b) end)\ncoroutine.resume(co, \"x\", \"y\")\n",
        "xy"
    },
    co_exh_yield_values => {
        "local co = coroutine.create(function() coroutine.yield(\"a\", \"b\") end)\nlocal _, r1, r2 = coroutine.resume(co)\nprint(r1, r2)\n",
        "a\tb"
    },
    co_exh_resume_yield_res => {
        "local co = coroutine.create(function() local x = coroutine.yield() print(x) end)\ncoroutine.resume(co)\ncoroutine.resume(co, 99)\n",
        "99"
    },
    co_exh_wrap_callable => {
        "local f = coroutine.wrap(function(x) return x + 1 end)\nprint(f(10))\n",
        "11"
    },
    co_exh_error_recovery => {
        "local co = coroutine.create(function() error(\"crash\") end)\nlocal ok, err = coroutine.resume(co)\nprint(ok, type(err))\n",
        "false\tstring"
    },
    co_exh_status_dead => {
        "local co = coroutine.create(function() end)\ncoroutine.resume(co)\nprint(coroutine.status(co))\n",
        "dead"
    },
    co_exh_yield_main_err => {
        "local ok = pcall(coroutine.yield)\nprint(ok)\n",
        "false"
    },
    co_exh_nested_yield => {
        "local function inner() coroutine.yield(\"inner\") end\nlocal co = coroutine.create(function() inner(); return \"outer\" end)\nlocal _, v1 = coroutine.resume(co)\nlocal _, v2 = coroutine.resume(co)\nprint(v1, v2)\n",
        "inner\touter"
    } }
