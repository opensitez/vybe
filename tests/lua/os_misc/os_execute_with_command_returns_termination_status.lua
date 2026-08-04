-- vybe-test: lua/os_misc/os_execute_with_command_returns_termination_status
-- origin: languages/lua/tests/lua/test_os_misc.rs

local __w1 = "true"
local __i = 0

local ok, status, code = os.execute("true")
do local __t = tostring(type(ok) == "boolean" or ok == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
