-- vybe-test: lua/base_rawset_insert/test_rawset_insert_trimmed
-- origin: languages/lua/tests/lua/test_base_rawset_insert.rs

local __w1 = "true"
local __i = 0

local t = {}
rawset(t, 3, 6)
do local __t = tostring(rawget(t, 3) == 6); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
