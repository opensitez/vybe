//! Functional programming patterns in pure Lua (Lua 5.x)

lua_print! {
    map_func => {
        "local function map(t, f)\n  local r = {}\n  for i, v in ipairs(t) do r[i] = f(v) end\n  return r\nend\nlocal r = map({1,2,3,4}, function(x) return x * x end)\nprint(r[1] .. \",\" .. r[4])\n",
        "1,16"
    },
    filter_func => {
        "local function filter(t, pred)\n  local r = {}\n  for _, v in ipairs(t) do\n    if pred(v) then r[#r+1] = v end\n  end\n  return r\nend\nlocal evens = filter({1,2,3,4,5,6}, function(x) return x % 2 == 0 end)\nprint(#evens .. \",\" .. evens[1])\n",
        "3,2"
    },
    reduce_func => {
        "local function reduce(t, f, init)\n  local acc = init\n  for _, v in ipairs(t) do acc = f(acc, v) end\n  return acc\nend\nprint(reduce({1,2,3,4,5}, function(a, b) return a + b end, 0))\n",
        "15"
    },
    compose_func => {
        "local function compose(f, g) return function(x) return f(g(x)) end end\nlocal double = function(x) return x * 2 end\nlocal inc = function(x) return x + 1 end\nlocal double_then_inc = compose(inc, double)\nprint(double_then_inc(5))\n",
        "11"
    },
    memoize_func => {
        "local function memoize(f)\n  local cache = {}\n  return function(n)\n    if cache[n] == nil then cache[n] = f(n) end\n    return cache[n]\n  end\nend\nlocal calls = 0\nlocal slow = memoize(function(n) calls = calls + 1; return n * n end)\nslow(5); slow(5); slow(5)\nprint(calls)\n",
        "1"
    },
    curry_func => {
        "local function curry(f) return function(a) return function(b) return f(a, b) end end end\nlocal add = curry(function(a, b) return a + b end)\nlocal add5 = add(5)\nprint(add5(3) .. \",\" .. add5(10))\n",
        "8,15"
    },
    partial_func => {
        "local function partial(f, ...)\n  local args = {...}\n  return function(...)\n    local all = {table.unpack(args)}\n    for _, v in ipairs({...}) do all[#all+1] = v end\n    return f(table.unpack(all))\n  end\nend\nlocal function mul(a, b) return a * b end\nlocal triple = partial(mul, 3)\nprint(triple(7))\n",
        "21"
    },
    pipeline_func => {
        "local function pipe(...)\n  local fns = {...}\n  return function(x)\n    for _, f in ipairs(fns) do x = f(x) end\n    return x\n  end\nend\nlocal process = pipe(\n  function(x) return x + 1 end,\n  function(x) return x * 2 end,\n  function(x) return x - 3 end\n)\nprint(process(5))\n",
        "9"
    },
    flatmap_func => {
        "local function flatmap(t, f)\n  local r = {}\n  for _, v in ipairs(t) do\n    for _, w in ipairs(f(v)) do r[#r+1] = w end\n  end\n  return r\nend\nlocal r = flatmap({1,2,3}, function(x) return {x, x*10} end)\nprint(table.concat(r, \",\"))\n",
        "1,10,2,20,3,30"
    },
    zip_func => {
        "local function zip(a, b)\n  local r = {}\n  for i = 1, math.min(#a, #b) do r[i] = {a[i], b[i]} end\n  return r\nend\nlocal z = zip({1,2,3}, {\"a\",\"b\",\"c\"})\nprint(z[2][1] .. z[2][2])\n",
        "2b"
    },
    flip_swaps_function_arguments => {
        "local function flip(f) return function(a, b) return f(b, a) end end\nlocal sub = function(a, b) return a - b end\nlocal rsub = flip(sub)\nprint(rsub(3, 10))\n",
        "7"
    },
    group_by_partitions_into_table_of_lists => {
        "local function group_by(t, key_fn)\n  local groups = {}\n  for _, v in ipairs(t) do\n    local k = key_fn(v)\n    if not groups[k] then groups[k] = {} end\n    groups[k][#groups[k]+1] = v\n  end\n  return groups\nend\nlocal g = group_by({1,2,3,4,5,6}, function(x) return x % 2 == 0 and 'even' or 'odd' end)\nprint(#g.even .. ',' .. #g.odd)\n",
        "3,3"
    },
    take_while_collects_prefix_matching_predicate => {
        "local function take_while(t, pred)\n  local r = {}\n  for _, v in ipairs(t) do\n    if not pred(v) then break end\n    r[#r+1] = v\n  end\n  return r\nend\nlocal t = take_while({1,2,3,4,5}, function(x) return x < 4 end)\nprint(table.concat(t, ','))\n",
        "1,2,3"
    },
    any_returns_true_if_predicate_matches_at_least_once => {
        "local function any(t, pred)\n  for _, v in ipairs(t) do if pred(v) then return true end end\n  return false\nend\nprint(tostring(any({1,3,5,4}, function(x) return x % 2 == 0 end)))\n",
        "true"
    },
    all_returns_false_if_one_value_fails_predicate => {
        "local function all(t, pred)\n  for _, v in ipairs(t) do if not pred(v) then return false end end\n  return true\nend\nprint(tostring(all({2,4,6,7}, function(x) return x % 2 == 0 end)))\n",
        "false"
    },
    scan_produces_running_accumulation => {
        "local function scan(t, f, init)\n  local r = {init}\n  for _, v in ipairs(t) do\n    r[#r+1] = f(r[#r], v)\n  end\n  return r\nend\nlocal s = scan({1,2,3,4}, function(a, b) return a + b end, 0)\nprint(table.concat(s, ','))\n",
        "0,1,3,6,10"
    },
    drop_while_removes_prefix_matching_predicate => {
        "local function drop_while(t, pred)\n  local i = 1\n  while i <= #t and pred(t[i]) do i = i + 1 end\n  local r = {}\n  for j = i, #t do r[#r+1] = t[j] end\n  return r\nend\nlocal t = drop_while({1,2,3,4,5}, function(x) return x < 3 end)\nprint(table.concat(t, ','))\n",
        "3,4,5"
    },
    function_identity_returns_its_argument => {
        "local function identity(x) return x end\nlocal val = {key = 'ok'}\nprint(identity(val).key)\n",
        "ok"
    } }
