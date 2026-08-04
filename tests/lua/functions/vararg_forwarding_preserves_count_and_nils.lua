-- vybe-test: lua/functions/vararg_forwarding_preserves_count_and_nils
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "10,nil,30"
local __i = 0

local function wrapper(...)
  return ...
end
local a, b, c = wrapper(10, nil, 30)
do local __t = tostring(tostring(a) .. ',' .. tostring(b) .. ',' .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
