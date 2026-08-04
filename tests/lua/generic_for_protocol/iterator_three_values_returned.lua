-- vybe-test: lua/generic_for_protocol/iterator_three_values_returned
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "3,z"
local __i = 0

local function iter(t, i)
  i = i + 1
  local v = t[i]
  if v ~= nil then return i, v end
end
local last_i, last_v
for i, v in iter, {"x", "y", "z"}, 0 do last_i = i; last_v = v end
do local __t = tostring(last_i .. "," .. last_v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
