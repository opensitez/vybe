-- vybe-test: lua/modules_require/test_require_not_found_error
-- origin: languages/lua/tests/lua/test_modules_require.rs

local __w1 = "false true"
local __i = 0

local ok, err = pcall(function() require('nonexistent_module') end); do local __t = tostring(tostring(ok)..' '..tostring(string.find(err, 'module') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
