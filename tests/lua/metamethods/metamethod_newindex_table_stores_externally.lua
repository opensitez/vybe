-- vybe-test: lua/metamethods/metamethod_newindex_table_stores_externally
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "v"
local __i = 0

local store={}
local t=setmetatable({}, {__newindex=store})
t.k="v"
do local __t = tostring(store.k); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
