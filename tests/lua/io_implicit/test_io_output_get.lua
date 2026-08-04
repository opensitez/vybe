-- vybe-test: lua/io_implicit/test_io_output_get
-- origin: languages/lua/tests/lua/test_io_implicit.rs

local __w1 = "file"
local __i = 0

do local __t = tostring(io.type(io.output())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
