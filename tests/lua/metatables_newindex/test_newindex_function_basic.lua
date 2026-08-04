-- vybe-test: lua/metatables_newindex/test_newindex_function_basic
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "a 10 true"
local __i = 0

local target, key, val; local t={}; setmetatable(t, {__newindex=function(tbl, k, v) target=tbl; key=k; val=v end}); t.a=10; do local __t = tostring(key..' '..val..' '..tostring(t==target)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
