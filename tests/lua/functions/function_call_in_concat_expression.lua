-- vybe-test: lua/functions/function_call_in_concat_expression
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "lang:lua"
local __i = 0

function name() return "lua" end
do local __t = tostring("lang:" .. name()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
