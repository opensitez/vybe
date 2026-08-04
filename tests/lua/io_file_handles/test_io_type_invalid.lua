-- vybe-test: lua/io_file_handles/test_io_type_invalid
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(io.type('not a file') or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
