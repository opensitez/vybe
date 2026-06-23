//! Coroutines — `coroutine.create`, `resume`, `yield`, `status` (Lua 5.x manual §2.6, §3.2).

lua_print! {
    coroutine_create_starts_suspended => {
        "local co = coroutine.create(function() end)\nprint(coroutine.status(co))\n",
        "suspended"
    },
    coroutine_resume_runs_function_body => {
        "local co = coroutine.create(function() return 7 end)\nlocal ok, v = coroutine.resume(co)\nprint(v)\n",
        "7"
    },
    coroutine_resume_returns_success_flag => {
        "local co = coroutine.create(function() return 1 end)\nlocal ok = coroutine.resume(co)\nprint(ok)\n",
        "true"
    },
    coroutine_yield_passes_value_to_resumer => {
        "local co = coroutine.create(function()\n  coroutine.yield(10)\n  return 20\nend)\nlocal _, a = coroutine.resume(co)\nlocal _, b = coroutine.resume(co)\nprint(a .. \",\" .. b)\n",
        "10,20"
    },
    coroutine_status_dead_after_completion => {
        "local co = coroutine.create(function() end)\ncoroutine.resume(co)\nprint(coroutine.status(co))\n",
        "dead"
    },
    coroutine_resume_passes_arguments_to_function => {
        "local co = coroutine.create(function(a, b) return a + b end)\nlocal _, v = coroutine.resume(co, 3, 4)\nprint(v)\n",
        "7"
    },
    coroutine_yield_receives_resume_arguments => {
        "local co = coroutine.create(function()\n  local x = coroutine.yield()\n  return x\nend)\ncoroutine.resume(co)\nlocal _, v = coroutine.resume(co, 99)\nprint(v)\n",
        "99"
    },
    coroutine_wrap_calls_function_directly => {
        "local f = coroutine.wrap(function() return \"ok\" end)\nprint(f())\n",
        "ok"
    },
    coroutine_running_is_nil_outside_coroutine => {
        "print(tostring(coroutine.running()))\n",
        "nil"
    },
    coroutine_isyieldable_outside_coroutine => {
        "print(coroutine.isyieldable())\n",
        "false"
    },
    coroutine_resume_returns_false_on_error => {
        "local co = coroutine.create(function() error(\"fail\") end)\nlocal ok = coroutine.resume(co)\nprint(ok)\n",
        "false"
    },
    coroutine_yield_from_nested_call => {
        "local function inner() coroutine.yield(5) end\nlocal co = coroutine.create(function() inner() return 1 end)\nlocal _, v = coroutine.resume(co)\nprint(v)\n",
        "5"
    },
    coroutine_running_inside_coroutine_is_thread => {
        "local co = coroutine.create(function() print(type(coroutine.running())) end)\ncoroutine.resume(co)\n",
        "thread"
    },
    coroutine_wrap_propagates_yield_values => {
        "local f = coroutine.wrap(function() coroutine.yield(3) return 9 end)\nprint(f())\n",
        "3"
    },
    coroutine_status_running_during_body => {
        "local seen = \"\"\nlocal co = coroutine.create(function()\n  seen = coroutine.status(coroutine.running())\nend)\ncoroutine.resume(co)\nprint(seen)\n",
        "running"
    },
    coroutine_resume_multiple_yields_in_sequence => {
        "local co = coroutine.create(function()\n  coroutine.yield(1)\n  coroutine.yield(2)\nend)\ncoroutine.resume(co)\nlocal _, a = coroutine.resume(co)\nprint(a)\n",
        "2"
    },
    coroutine_dead_after_error_in_body => {
        "local co = coroutine.create(function() error(\"x\") end)\npcall(coroutine.resume, co)\nprint(coroutine.status(co))\n",
        "dead"
    },
    coroutine_yield_passes_multiple_return_values => {
        "local co = coroutine.create(function() return 1, 2, 3 end)\nlocal _, a, b, c = coroutine.resume(co)\nprint(a + b + c)\n",
        "6"
    },
    coroutine_create_requires_function_argument => {
        "local ok = pcall(coroutine.create, 1)\nprint(ok)\n",
        "false"
    },
}
