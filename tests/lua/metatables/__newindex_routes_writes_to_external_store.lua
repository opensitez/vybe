-- vybe-test: lua/metatables/__newindex_routes_writes_to_external_store
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "1"
local __i = 0

local data = {}
local t = setmetatable({}, {__newindex = data})
t.a = 1
do local __t = tostring(data.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
