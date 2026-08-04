-- vybe-test: lua/generic_for_protocol/iterator_explicit_state
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "60"
local __i = 0

local function iter(t, i)
  i = i + 1
  local v = t[i]
  if v then return i, v end
end
local s = 0
for i, v in iter, {10, 20, 30}, 0 do s = s + v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
