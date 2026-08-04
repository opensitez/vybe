-- vybe-test: lua/pcall_nested_xpcall_scenarios/xpcall_custom_class
-- origin: languages/lua/tests/lua/test_pcall_nested_xpcall_scenarios.rs

local __w1 = "false\tbad\t500"
local __i = 0

local function handler(e) return {msg = e.msg, code = e.code} end
local ok, err = xpcall(function() error({msg="bad", code=500}) end, handler)
do local __t = tostring(ok) .. "\t" .. tostring(err.msg) .. "\t" .. tostring(err.code); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
