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
    collectgarbage_default_performs_full_collection => {
        "local ok = pcall(collectgarbage)\nprint(ok)\n",
        "true"
    },
    collectgarbage_collect_action_performs_full_collection => {
        "local ok = pcall(collectgarbage, \"collect\")\nprint(ok)\n",
        "true"
    },
    collectgarbage_setpause_modifies_pause_and_returns_previous => {
        "local prev = collectgarbage(\"setpause\", 150)\nlocal cur = collectgarbage(\"setpause\", prev)\nprint(type(prev) == \"number\" and cur == 150)\n",
        "true"
    },
    collectgarbage_setstepmul_modifies_multiplier_and_returns_previous => {
        "local prev = collectgarbage(\"setstepmul\", 250)\nlocal cur = collectgarbage(\"setstepmul\", prev)\nprint(type(prev) == \"number\" and cur == 250)\n",
        "true"
    },
    collectgarbage_invalid_action_raises_error => {
        "local ok, err = pcall(collectgarbage, \"invalid_garbage_action\")\nprint(ok)\n",
        "false"
    },
    collectgarbage_step_with_large_size_finishes_cycle => {
        "local finished = collectgarbage(\"step\", 1000000)\nprint(type(finished) == \"boolean\")\n",
        "true"
    },
    collectgarbage_garbage_creation_increases_memory_count => {
        "local before = collectgarbage(\"count\")\nlocal t = {}\nfor i=1,1000 do t[i] = {x=i} end\nlocal after = collectgarbage(\"count\")\nprint(after > before)\n",
        "true"
    },
    collectgarbage_collect_frees_temporary_garbage => {
        "local before = collectgarbage(\"count\")\nlocal function make_garbage()\n  local t = {}\n  for i=1,1000 do t[i] = {x=i} end\nend\nmake_garbage()\ncollectgarbage(\"collect\")\nlocal after = collectgarbage(\"count\")\nprint(after - before < 100)\n",
        "true"
    },
}
