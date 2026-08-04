-- vybe-test: lua/weak_tables/weak_both_mode_kv
-- origin: languages/lua/tests/lua/test_weak_tables.rs

local __w1 = "kv"
local __i = 0

local t = setmetatable({}, {__mode = "kv"})
do local __t = tostring(getmetatable(t).__mode); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
