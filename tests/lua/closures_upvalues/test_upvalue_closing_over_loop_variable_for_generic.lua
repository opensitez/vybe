-- vybe-test: lua/closures_upvalues/test_upvalue_closing_over_loop_variable_for_generic
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "ab"
local __i = 0

local t={}; for _, v in ipairs({'a','b'}) do table.insert(t, function() return v end) end; do local __t = tostring(t[1]()..t[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
