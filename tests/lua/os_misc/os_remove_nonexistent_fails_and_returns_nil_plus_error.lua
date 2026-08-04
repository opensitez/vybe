-- vybe-test: lua/os_misc/os_remove_nonexistent_fails_and_returns_nil_plus_error
-- origin: languages/lua/tests/lua/test_os_misc.rs

local __w1 = "true"
local __i = 0

local ok, err = os.remove("nonexistent_file_xyz_123.txt")
do local __t = tostring(ok == nil and type(err) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
