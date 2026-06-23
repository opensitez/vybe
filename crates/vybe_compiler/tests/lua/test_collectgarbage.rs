//! Garbage collection — `collectgarbage` (Lua 5.x manual §2.5.1).

lua_print! {
    collectgarbage_count_returns_kilobytes_number => {
        "print(type(collectgarbage(\"count\")) == \"number\")\n",
        "true"
    },
    collectgarbage_isrunning_returns_boolean => {
        "print(type(collectgarbage(\"isrunning\")) == \"boolean\")\n",
        "true"
    },
    collectgarbage_restart_after_stop => {
        "collectgarbage(\"stop\")\nlocal was = collectgarbage(\"isrunning\")\ncollectgarbage(\"restart\")\nprint(was == false and collectgarbage(\"isrunning\"))\n",
        "true"
    },
    collectgarbage_step_advances_collector => {
        "local ok = pcall(function() collectgarbage(\"step\", 0) end)\nprint(ok)\n",
        "true"
    },
}
