-- vybe-test: lua/functional_patterns/pipeline_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "9"
local __i = 0

local function pipe(...)
  local fns = {...}
  return function(x)
    for _, f in ipairs(fns) do x = f(x) end
    return x
  end
end
local process = pipe(
  function(x) return x + 1 end,
  function(x) return x * 2 end,
  function(x) return x - 3 end
)
do local __t = tostring(process(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
