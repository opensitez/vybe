-- App 06: Paginated Feed Reader (iterator-backed stream processing)

local pages = {
  {
    {id = 1, title = "Release notes", author = "team"},
    {id = 2, title = "Incident postmortem", author = "sre"},
  },
  {
    {id = 3, title = "Roadmap", author = "pm"},
    {id = 4, title = "Schema migration", author = "db"},
  }
}

local function feed_iter(dataset)
  local page_i, item_i = 1, 0
  return function()
    while dataset[page_i] do
      item_i = item_i + 1
      local item = dataset[page_i][item_i]
      if item then return item end
      page_i, item_i = page_i + 1, 0
    end
  end
end

local counts = {}
for post in feed_iter(pages) do
  counts[post.author] = (counts[post.author] or 0) + 1
  print(string.format("[%d] %s (%s)", post.id, post.title, post.author))
end

print("posts by author")
for author, n in pairs(counts) do
  print(author, n)
end
