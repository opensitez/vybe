-- vybe-test: lua/metatables_index_chains/newindex_missing_only
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "1 99 10"
local __i = 0

local called = 0
local t = setmetatable({x = 1}, {
  __newindex = function(tbl, k, v)
    called = called + 1
    rawset(tbl, k, v)
  end
})
t.x = 99   -- existing, no __newindex
t.y = 10   -- new, triggers
do local __t = tostring(called) .. "\t" .. tostring(t.x) .. "\t" .. tostring(t.y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
