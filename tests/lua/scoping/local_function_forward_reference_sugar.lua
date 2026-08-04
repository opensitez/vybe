-- vybe-test: lua/scoping/local_function_forward_reference_sugar
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "24"
local __i = 0

local function fact(n)
  if n <= 1 then return 1 end
  return n * fact(n - 1)
end
do local __t = tostring(fact(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
