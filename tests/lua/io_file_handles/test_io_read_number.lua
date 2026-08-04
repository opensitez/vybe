-- vybe-test: lua/io_file_handles/test_io_read_number
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local __w1 = "123.45"
local __i = 0

local f = io.tmpfile(); f:write('123.45'); f:seek('set'); do local __t = tostring(f:read('n')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
