-- vybe-test: lua/raw_access/rawget_missing_nil
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "nil"
local __i = 0

local t = setmetatable({}, {__index = function() return 99 end})
do local __t = tostring(tostring(rawget(t, "missing"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
