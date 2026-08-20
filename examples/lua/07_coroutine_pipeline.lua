-- App 07: ETL Pipeline (coroutine producer/transformer/sink)

local input = {
  "id=1,name=alice,score=81",
  "id=2,name=bob,score=42",
  "id=3,name=claire,score=93",
}

local function producer(rows)
  return coroutine.create(function()
    for _, row in ipairs(rows) do coroutine.yield(row) end
  end)
end

local function parser(src)
  return coroutine.create(function()
    while true do
      local ok, row = coroutine.resume(src)
      if not ok or not row then break end
      local rec = {}
      for k, v in row:gmatch("([%a_]+)=([%w_]+)") do
        rec[k] = tonumber(v) or v
      end
      coroutine.yield(rec)
    end
  end)
end

local function grader(src)
  return coroutine.create(function()
    while true do
      local ok, rec = coroutine.resume(src)
      if not ok or not rec then break end
      rec.grade = rec.score >= 90 and "A" or (rec.score >= 70 and "B" or "C")
      coroutine.yield(rec)
    end
  end)
end

local src = producer(input)
local parsed = parser(src)
local graded = grader(parsed)

while true do
  local ok, rec = coroutine.resume(graded)
  if not ok or not rec then break end
  print(string.format("id=%d name=%s score=%d grade=%s", rec.id, rec.name, rec.score, rec.grade))
end
