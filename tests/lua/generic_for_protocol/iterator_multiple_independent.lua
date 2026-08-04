-- vybe-test: lua/generic_for_protocol/iterator_multiple_independent
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "3,3"
local __i = 0

local function counter()
  local n = 0
  return function() n = n + 1; return n <= 2 and n or nil end
end
local a, b = 0, 0
for v in counter() do a = a + v end
for v in counter() do b = b + v end
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
