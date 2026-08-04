-- vybe-test: lua/scoping/local_declaration_hides_same_name_global
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "local\nglobal"
local __i = 0

g_conflict = 'global'
do
  local g_conflict = 'local'
  do local __t = tostring(g_conflict); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(g_conflict); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
