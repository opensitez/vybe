-- vybe-test: lua/io_library/io_type_on_closed_file
-- origin: languages/lua/tests/lua/test_io_library.rs

local __w1 = "true"
local __i = 0

local f = io.tmpfile()
if f then f:close() do local __t = tostring(io.type(f) == "closed file" or io.type(f) == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring(true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
