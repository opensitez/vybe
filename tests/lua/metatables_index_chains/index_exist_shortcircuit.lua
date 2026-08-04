-- vybe-test: lua/metatables_index_chains/index_exist_shortcircuit
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "99 false"
local __i = 0

local called = false
local t = setmetatable({k = 99}, {
  __index = function() called = true; return 0 end
})
local v = t.k
do local __t = tostring(v) .. "\t" .. tostring(called); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
