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

local function make_caption_from_paras(paras)
  local caption_blocks = {}
  for _, para in ipairs(paras) do
    caption_blocks[#caption_blocks + 1] = pandoc.Plain(para.content)
  end

  return {
    long = caption_blocks,
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
      and current.t == "Para"
      and is_table_caption_text(stringify_inlines(current.content))
    then
      local caption_paras = { current }
      local j = i + 1

      while j <= #blocks
        and blocks[j].t == "Para"
        and is_table_caption_text(stringify_inlines(blocks[j].content))
      do
        caption_paras[#caption_paras + 1] = blocks[j]
        j = j + 1
      end

      if j <= #blocks and blocks[j].t == "Table" then
        blocks[j].caption = make_caption_from_paras(caption_paras)
        out[#out + 1] = blocks[j]
        i = j + 1
      else
        out[#out + 1] = current
        i = i + 1
      end
    else
      out[#out + 1] = current
      i = i + 1
    end
  end

  return pandoc.Pandoc(out, doc.meta)
end

return { Pandoc = Pandoc }
