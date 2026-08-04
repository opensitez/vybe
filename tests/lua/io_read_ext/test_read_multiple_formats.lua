-- vybe-test: lua/io_read_ext/test_read_multiple_formats
-- origin: languages/lua/tests/lua/test_io_read_ext.rs

local __w1 = "123  abc xyz"
local __i = 0

local f = io.tmpfile(); f:write('123 abc\nxyz'); f:seek('set'); local n, w, l = f:read('n', 4, 'l'); do local __t = tostring(n..' '..w..' '..l); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
