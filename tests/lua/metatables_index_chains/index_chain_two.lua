-- vybe-test: lua/metatables_index_chains/index_chain_two
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "10"
local __i = 0

local base = {x = 10}
local mid = setmetatable({}, {__index = base})
local top = setmetatable({}, {__index = mid})
do local __t = tostring(top.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
