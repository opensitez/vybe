-- vybe-test: lua/raw_access/rawlen_ignores_len
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "3"
local __i = 0

local t = setmetatable({1, 2, 3}, {__len = function() return 999 end})
do local __t = tostring(rawlen(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
