-- vybe-test: lua/raw_access/rawset_no_newindex_trigger
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "false"
local __i = 0

local called = false
local t = setmetatable({}, {__newindex = function() called = true end})
rawset(t, "x", 1)
do local __t = tostring(called); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
