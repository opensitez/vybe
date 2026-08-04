-- vybe-test: lua/coroutine_status_initial/test_status_with_nested_resume
-- origin: languages/lua/tests/lua/test_coroutine_status_initial.rs

local __w1 = "true"
local __i = 0

local outer
outer = coroutine.create(function()
  local ok, innerStatus = pcall(function() return coroutine.status(inner) end)
  return ok and innerStatus == nil
end)
local inner = coroutine.create(function() return 1 end)
coroutine.resume(outer)
do local __t = tostring(1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end")#), "1");
}

#[test]
fn test_status_for_two_depth() {
    assert_eq!(run_lua_one(r#"local t1 = coroutine.create(function() return 1 end)
local t2 = coroutine.create(function() return coroutine.resume(t1) end)
coroutine.resume(t2)
do local __t = tostring(coroutine.status(t2) == "dead"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
