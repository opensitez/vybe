-- vybe-test: lua/metatables/__index_function_computes_missing_key
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "3"
local __i = 0

local t=setmetatable({}, {__index=function(tbl,k) return #k end})
do local __t = tostring(t.abc); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
