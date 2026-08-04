-- vybe-test: lua/metatables_index_chains/newindex_proxy
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "10"
local __i = 0

local store = {}
local proxy = setmetatable({}, {
  __newindex = function(_, k, v) store[k] = v * 2 end,
  __index = store
})
proxy.x = 5
do local __t = tostring(proxy.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
