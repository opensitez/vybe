-- vybe-test: lua/ipairs_pairs_edge/ipairs_one_based_indices
-- origin: languages/lua/tests/lua/test_ipairs_pairs_edge.rs

local __w1 = "1"
local __i = 0

local first_i
for i, _ in ipairs({"a", "b"}) do
  if not first_i then first_i = i end
end
do local __t = tostring(first_i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
