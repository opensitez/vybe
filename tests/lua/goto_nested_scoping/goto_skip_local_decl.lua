-- vybe-test: lua/goto_nested_scoping/goto_skip_local_decl
-- origin: languages/lua/tests/lua/test_goto_nested_scoping.rs

local __w1 = "true"
local __i = 0

local ok = true
goto target
local val = 42
::target::
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
