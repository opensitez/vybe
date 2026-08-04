-- vybe-test: lua/io_read_ext/test_read_zero_bytes
-- origin: languages/lua/tests/lua/test_io_read_ext.rs

local __w1 = ""
local __i = 0

local f = io.tmpfile(); do local __t = tostring(f:read(0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
