-- vybe-test: lua/raw_access/rawequal_identity
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "true"
local __i = 0

local t = {}
do local __t = tostring(rawequal(t, t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
