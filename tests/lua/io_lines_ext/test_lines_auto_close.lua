-- vybe-test: lua/io_lines_ext/test_lines_auto_close
-- origin: languages/lua/tests/lua/test_io_lines_ext.rs

local __w1 = "abc"
local __i = 0

local n = os.tmpname(); local f = io.open(n, 'w'); f:write('abc'); f:close(); local r=''; for l in io.lines(n) do r=r..l end; os.remove(n); do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
