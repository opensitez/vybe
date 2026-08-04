-- vybe-test: lua/weak_tables/weak_keys_mode_recorded_in_metatable
-- origin: languages/lua/tests/lua/test_weak_tables.rs

local __w1 = "k"
local __i = 0

local m = {__mode = "k"}
local t = setmetatable({}, m)
do local __t = tostring(getmetatable(t).__mode); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
