-- vybe-test: lua/programs/lexical_scope_shadows_global_name
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

x = 1
local function f()
  local x = 2
  return x
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
