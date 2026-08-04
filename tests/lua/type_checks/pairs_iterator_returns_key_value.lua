-- vybe-test: lua/type_checks/pairs_iterator_returns_key_value
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "a=1"
local __i = 0

local t={a=1}
local k,v=next(t)
do local __t = tostring(k.."="..v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
