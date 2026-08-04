-- vybe-test: lua/generic_for_protocol/custom_iterator_func
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "15"
local __i = 0

local function range(n)
  local i = 0
  return function()
    i = i + 1
    if i <= n then return i end
  end
end
local s = 0
for v in range(5) do s = s + v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
