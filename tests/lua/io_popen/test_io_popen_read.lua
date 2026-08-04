-- vybe-test: lua/io_popen/test_io_popen_read
-- origin: languages/lua/tests/lua/test_io_popen.rs

local __w1 = "string"
local __i = 0

local f = io.popen('echo test'); local r = f:read('*a'); f:close(); do local __t = tostring(type(r)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
