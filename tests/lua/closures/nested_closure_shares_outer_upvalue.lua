-- vybe-test: lua/closures/nested_closure_shares_outer_upvalue
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "1"
local __i = 0

local n=0
local function outer()
  local function inner() n=n+1 end
  return inner
end
local inc=outer()
inc()
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
