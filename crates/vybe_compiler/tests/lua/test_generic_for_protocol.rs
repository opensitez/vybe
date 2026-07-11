//! Generic `for` loop protocol — iterator functions and state propagation (Lua 5.x §3.3.5)

lua_print! {
    custom_iterator_func => {
        "local function range(n)\n  local i = 0\n  return function()\n    i = i + 1\n    if i <= n then return i end\n  end\nend\nlocal s = 0\nfor v in range(5) do s = s + v end\nprint(s)\n",
        "15"
    },
    iterator_closure_state => {
        "local function values(t)\n  local i = 0\n  return function()\n    i = i + 1\n    return t[i]\n  end\nend\nlocal r = \"\"\nfor v in values({\"a\", \"b\", \"c\"}) do r = r .. v end\nprint(r)\n",
        "abc"
    },
    iterator_explicit_state => {
        "local function iter(t, i)\n  i = i + 1\n  local v = t[i]\n  if v then return i, v end\nend\nlocal s = 0\nfor i, v in iter, {10, 20, 30}, 0 do s = s + v end\nprint(s)\n",
        "60"
    },
    iterator_terminate_on_nil => {
        "local function upto3(t, i)\n  i = i + 1\n  if i <= 3 then return i end\nend\nlocal n = 0\nfor _ in upto3, {}, 0 do n = n + 1 end\nprint(n)\n",
        "3"
    },
    iterator_coroutine_wrapper => {
        "local function gen(n)\n  return coroutine.wrap(function()\n    for i = 1, n do coroutine.yield(i) end\n  end)\nend\nlocal s = 0\nfor v in gen(4) do s = s + v end\nprint(s)\n",
        "10"
    },
    iterator_multiple_independent => {
        "local function counter()\n  local n = 0\n  return function() n = n + 1; return n <= 2 and n or nil end\nend\nlocal a, b = 0, 0\nfor v in counter() do a = a + v end\nfor v in counter() do b = b + v end\nprint(a .. \",\" .. b)\n",
        "3,3"
    },
    iterator_three_values_returned => {
        "local function iter(t, i)\n  i = i + 1\n  local v = t[i]\n  if v ~= nil then return i, v end\nend\nlocal last_i, last_v\nfor i, v in iter, {\"x\", \"y\", \"z\"}, 0 do last_i = i; last_v = v end\nprint(last_i .. \",\" .. last_v)\n",
        "3,z"
    },
}
