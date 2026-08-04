-- vybe-test: lua/error_levels/test_error_level_2
-- origin: languages/lua/tests/lua/test_error_levels.rs

local __w1 = "false true"
local __i = 0

local function f() error('boom', 2) end; local ok, err = pcall(function() f() end); do local __t = tostring(tostring(ok)..' '..tostring(string.find(err, 'boom') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
