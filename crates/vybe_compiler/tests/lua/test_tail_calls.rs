//! Tail calls — proper tail recursion (Lua 5.x manual §3.4.11).

lua_print! {
    tail_call_returns_like_direct_call => {
        "local function tail(n)\n  if n == 0 then return \"ok\" end\n  return tail(n - 1)\nend\nprint(tail(3))\n",
        "ok"
    },
    tail_call_with_return_values => {
        "local function add(n, acc)\n  if n == 0 then return acc end\n  return add(n - 1, acc + n)\nend\nprint(add(5, 0))\n",
        "15"
    },
    non_tail_call_still_correct_without_optimization => {
        "local function not_tail(n)\n  if n == 0 then return 0 end\n  local x = not_tail(n - 1)\n  return x\nend\nprint(not_tail(4))\n",
        "0"
    },
    mutual_tail_calls_alternate => {
        "local even, odd\nfunction even(n) if n == 0 then return true end return odd(n - 1) end\nfunction odd(n) if n == 0 then return false end return even(n - 1) end\nprint(even(4))\n",
        "true"
    },
}
