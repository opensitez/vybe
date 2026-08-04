-- vybe-test: lua/io_lines_ext/test_lines_multiple_formats
-- origin: languages/lua/tests/lua/test_io_lines_ext.rs

local __w1 = "123  abc"
local __i = 0

local f = io.tmpfile(); f:write('123 abc\n'); f:seek('set'); local n, w; for a, b in f:lines('n', 'l') do n, w = a, b end; do local __t = tostring(n..' '..w); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
