-- vybe-test: lua/ipairs_pairs_edge/ipairs_single
-- origin: languages/lua/tests/lua/test_ipairs_pairs_edge.rs

local __w1 = "1=42"
local __i = 0

for i, v in ipairs({42}) do do local __t = tostring(i .. "=" .. v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
