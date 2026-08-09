//! pcall retry, fallback, and cleanup patterns (Lua 5.x §6.1)

lua_print! {
pcall_div_fallback => {
    "local function try_div(a, b)\n  if b == 0 then error(\"div0\") end\n  return a / b\nend\nlocal ok, v = pcall(try_div, 10, 0)\nlocal result = ok and v or -1\nprint(result)\n",
    "-1"
},
pcall_retry_loop => {
    "local attempts = 0\nlocal ok = false\nwhile not ok and attempts < 3 do\n  attempts = attempts + 1\n  ok = pcall(function()\n    if attempts < 3 then error(\"retry\") end\n  end)\nend\nprint(attempts)\n",
    "3"
},
pcall_error_cleanup => {
    "local cleaned = false\nlocal ok = pcall(function()\n  error(\"boom\")\nend)\ncleaned = true\nprint(cleaned)\n",
    "true"
},
pcall_nil_member => {
    "local ok, err = pcall(function()\n  local t = nil\n  return t.x\nend)\nprint(ok)\n",
    "false"
},
pcall_batch_dangerous => {
    "local function dangerous(n)\n  if n > 10 then error(\"too big\") end\n  return n * 2\nend\nlocal results = {}\nfor _, v in ipairs({3, 15, 7}) do\n  local ok, r = pcall(dangerous, v)\n  results[#results+1] = ok and r or \"err\"\nend\nprint(table.concat(results, \",\"))\n",
    "6,err,14"
},
pcall_varargs => {
    "local function sum(...)\n  local s = 0\n  for _, v in ipairs({...}) do s = s + v end\n  return s\nend\nlocal ok, v = pcall(sum, 1, 2, 3, 4)\nprint(ok, v)\n",
    "true\t10"
},
pcall_nested_unaffected => {
    "local outer_ok = pcall(function()\n  local inner_ok = pcall(function() error(\"inner\") end)\n  assert(inner_ok == false)\nend)\nprint(outer_ok)\n",
    "true"
},
pcall_returns_all => {
    "local function multi() return 1, 2, 3 end\nlocal ok, a, b, c = pcall(multi)\nprint(ok, a, b, c)\n",
    "true\t1\t2\t3"
} }
