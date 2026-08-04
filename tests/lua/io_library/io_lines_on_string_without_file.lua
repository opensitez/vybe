-- vybe-test: lua/io_library/io_lines_on_string_without_file
-- origin: languages/lua/tests/lua/test_io_library.rs

local __w1 = "6"
local __i = 0

local sum = 0
for n in io.lines("1\n2\n3\n") do sum = sum + tonumber(n) end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
