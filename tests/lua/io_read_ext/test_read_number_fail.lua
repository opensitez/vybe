-- vybe-test: lua/io_read_ext/test_read_number_fail
-- origin: languages/lua/tests/lua/test_io_read_ext.rs

local __w1 = "nil"
local __i = 0

local f = io.tmpfile(); f:write('abc'); f:seek('set'); do local __t = tostring(f:read('n') or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
