-- Promote a paragraph like "表1 标题" or "Table 1. Caption" immediately
-- preceding a Pandoc Table into that table's caption.

local function stringify_inlines(inlines)
  return pandoc.utils.stringify(inlines or {}):gsub("^%s+", ""):gsub("%s+$", "")
end

local function is_table_caption_text(text)
  if text == "" then
    return false
  end

  return text:match("^表%s*[%d一二三四五六七八九十百千]+[%s:：.、%-_].*")
      or text:match("^表%s*[%d一二三四五六七八九十百千]+$")
      or text:match("^Table%s*[%dIVXLCivxlc]+[%s:：.、%-_].*")
      or text:match("^Table%s*[%dIVXLCivxlc]+$")
end

local function make_caption_from_para(para)
  return {
    long = { pandoc.Plain(para.content) },
    short = nil,
  }
end

function Pandoc(doc)
  local blocks = doc.blocks
  local out = {}
  local i = 1

  while i <= #blocks do
    local current = blocks[i]
    local next_block = blocks[i + 1]

    if current
      and next_block
      and current.t == "Para"
      and next_block.t == "Table"
      and is_table_caption_text(stringify_inlines(current.content))
    then
      next_block.caption = make_caption_from_para(current)
      out[#out + 1] = next_block
      i = i + 2
    else
      out[#out + 1] = current
      i = i + 1
    end
  end

  return pandoc.Pandoc(out, doc.meta)
end

return { Pandoc = Pandoc }
