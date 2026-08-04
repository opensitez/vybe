-- vybe-test: lua/metatables/metatable_index_chain_three_levels
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "base"
local __i = 0

local base = {method = function() return 'base' end}
local mid = setmetatable({}, {__index = base})
local top = setmetatable({}, {__index = mid})
do local __t = tostring(top.method()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
