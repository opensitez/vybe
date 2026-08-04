-- vybe-test: lua/load/load_chunkname_for_debug_traceback
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "true"
local __i = 0

local f = load("error('test_err')", "custom_chunk_name")
local ok, err = pcall(f)
do local __t = tostring(err:match("custom_chunk_name") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
