-- vybe-test: lua/metatables/rawset_writes_without_triggering_newindex
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "0"
local __i = 0

local log = 0
local t = setmetatable({}, {__newindex = function() log = log + 1 end})
rawset(t, "k", 1)
do local __t = tostring(log); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
