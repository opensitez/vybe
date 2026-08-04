-- vybe-test: lua/io_file_handles/test_io_lines_iterator
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local __w1 = "ab"
local __i = 0

local f = io.tmpfile(); f:write('a\nb\n'); f:seek('set'); local s=''; for l in f:lines() do s=s..l end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
