-- vybe-test: lua/programs/comma_split_into_array_manual
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "a/b/c"
local __i = 0

local s = "a,b,c"
local t = {}
for part in string.gmatch(s, "[^,]+") do table.insert(t, part) end
do local __t = tostring(table.concat(t, "/")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
