-- 11_environment_and_sandbox.lua
-- Demonstrates _ENV, load with custom environments, and restricted execution.

local safe_env = {
  print = print,
  math = math,
  tonumber = tonumber,
  tostring = tostring
}

local code = [[
  local x = math.sqrt(144)
  print("sandbox sqrt:", x)
  return x * 2
]]

local chunk, err = load(code, "sandbox_chunk", "t", safe_env)
if not chunk then
  error(err)
end

local result = chunk()
print("sandbox result:", result)

local restricted, err2 = load("return os.date()", "restricted", "t", safe_env)
if not restricted then
  error(err2)
end

local ok, value = pcall(restricted)
print("restricted access ok:", ok)
print("restricted value:", value)
