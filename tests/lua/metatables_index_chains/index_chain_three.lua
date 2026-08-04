-- vybe-test: lua/metatables_index_chains/index_chain_three
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "42"
local __i = 0

local L1 = {v = 42}
local L2 = setmetatable({}, {__index = L1})
local L3 = setmetatable({}, {__index = L2})
local L4 = setmetatable({}, {__index = L3})
do local __t = tostring(L4.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
