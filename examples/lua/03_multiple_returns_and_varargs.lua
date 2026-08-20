-- App 03: Sensor Aggregator (varargs batches, multi-value reports)

local function batch_stats(batch_name, ...)
  local n = select("#", ...)
  local sum, minv, maxv = 0, math.huge, -math.huge
  for i = 1, n do
    local v = select(i, ...)
    sum = sum + v
    if v < minv then minv = v end
    if v > maxv then maxv = v end
  end
  local avg = n > 0 and (sum / n) or 0
  return batch_name, n, minv, maxv, avg
end

local function print_report(...)
  local batches = {...}
  for _, row in ipairs(batches) do
    local name, n, minv, maxv, avg = table.unpack(row)
    print(string.format("%s -> count=%d min=%.2f max=%.2f avg=%.2f", name, n, minv, maxv, avg))
  end
end

local reports = {
  {batch_stats("rack-a", 20.1, 19.8, 21.2, 20.9, 20.2)},
  {batch_stats("rack-b", 30.1, 29.7, 28.9, 31.0)},
  {batch_stats("rack-c", 15.0, 14.8, 15.4, 15.2, 15.1, 14.9)},
}

print_report(table.unpack(reports))
