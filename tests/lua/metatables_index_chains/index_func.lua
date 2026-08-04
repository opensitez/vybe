-- vybe-test: lua/metatables_index_chains/index_func
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "HELLO"
local __i = 0

local t = setmetatable({}, {
  __index = function(_, k) return k:upper() end
})
do local __t = tostring(t.hello); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
