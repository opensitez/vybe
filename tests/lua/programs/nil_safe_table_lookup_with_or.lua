-- vybe-test: lua/programs/nil_safe_table_lookup_with_or
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "default"
local __i = 0

local t = {}
local v = t.missing or "default"
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
