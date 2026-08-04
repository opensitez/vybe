-- vybe-test: lua/base_load_chunk/test_load_chunk_edge_last
-- origin: languages/lua/tests/lua/test_base_load_chunk.rs

local __w1 = "true"
local __i = 0

local f = assert(load("return 18+19")); do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
