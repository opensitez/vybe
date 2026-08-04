-- vybe-test: lua/table_constructors/constructor_nil_key_entry_errors
-- origin: languages/lua/tests/lua/test_table_constructors.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() return {[nil]=1, a=2} end)
do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
