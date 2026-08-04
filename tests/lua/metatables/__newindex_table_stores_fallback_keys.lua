-- vybe-test: lua/metatables/__newindex_table_stores_fallback_keys
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "1"
local __i = 0

local store={}
local t=setmetatable({},{__newindex=store})
t.x=1
do local __t = tostring(store.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
