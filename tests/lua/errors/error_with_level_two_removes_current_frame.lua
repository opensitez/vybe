-- vybe-test: lua/errors/error_with_level_two_removes_current_frame
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "true"
local __i = 0

local function f() error("my_err", 2) end
local ok, msg = pcall(f)
do local __t = tostring(type(msg) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
