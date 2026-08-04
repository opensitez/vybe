-- vybe-test: lua/ipairs_pairs_edge/ipairs_iterator_tuple
-- origin: languages/lua/tests/lua/test_ipairs_pairs_edge.rs

local __w1 = "10"
local __i = 0

local it, s, i = ipairs({10, 20, 30})
local _, v = it(s, i)
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
