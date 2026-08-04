-- vybe-test: lua/language_semantics_extended/assign_evaluate_first
-- origin: languages/lua/tests/lua/test_language_semantics_extended.rs

local __w1 = "2,99,nil"
local __i = 0

local t = {10}
local i = 1
i, t[i] = 2, 99
do local __t = tostring(i .. "," .. t[1] .. "," .. tostring(t[2])); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
