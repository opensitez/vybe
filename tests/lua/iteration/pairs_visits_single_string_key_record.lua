-- vybe-test: lua/iteration/pairs_visits_single_string_key_record
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "mode"
local __i = 0

local t = {mode = "rw"}
for k in pairs(t) do do local __t = tostring(k); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
