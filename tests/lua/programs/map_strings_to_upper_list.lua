-- vybe-test: lua/programs/map_strings_to_upper_list
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "AB"
local __i = 0

local t = {"a", "b"}
for i = 1, #t do t[i] = string.upper(t[i]) end
do local __t = tostring(table.concat(t, "")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
