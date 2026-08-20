-- App 38: Build Pipeline Simulator

local stages = {
  {name = "lint", duration = 2, fail_rate = 0.05},
  {name = "test", duration = 5, fail_rate = 0.15},
  {name = "package", duration = 1, fail_rate = 0.02},
  {name = "deploy", duration = 3, fail_rate = 0.10},
}

local function run_pipeline(commit)
  local elapsed = 0
  for _, s in ipairs(stages) do
    elapsed = elapsed + s.duration
    if math.random() < s.fail_rate then
      return false, s.name, elapsed
    end
  end
  return true, "done", elapsed
end

for i = 1, 5 do
  local ok, stage, t = run_pipeline("c" .. i)
  print(string.format("commit c%d -> %s at %s (%ds)", i, ok and "SUCCESS" or "FAILED", stage, t))
end
