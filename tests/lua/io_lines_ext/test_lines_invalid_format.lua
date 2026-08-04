-- vybe-test: lua/io_lines_ext/test_lines_invalid_format
-- origin: languages/lua/tests/lua/test_io_lines_ext.rs

local __w1 = "false"
local __i = 0

local f = io.tmpfile(); local ok = pcall(function() for l in f:lines('invalid') do end end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
