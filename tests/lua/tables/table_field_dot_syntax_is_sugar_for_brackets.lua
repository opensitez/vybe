-- vybe-test: lua/tables/table_field_dot_syntax_is_sugar_for_brackets
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "lua"
local __i = 0

local t = {name = "lua"}
do local __t = tostring(t["name"]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
