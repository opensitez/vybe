-- vybe-test: lua/functional_patterns/function_identity_returns_its_argument
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "ok"
local __i = 0

local function identity(x) return x end
local val = {key = 'ok'}
do local __t = tostring(identity(val).key); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
