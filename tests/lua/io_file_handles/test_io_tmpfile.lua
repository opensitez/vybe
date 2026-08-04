-- vybe-test: lua/io_file_handles/test_io_tmpfile
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local __w1 = "true"
local __i = 0

local f = io.tmpfile(); do local __t = tostring(type(f) == 'userdata'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
