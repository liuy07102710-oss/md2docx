-- Promote a paragraph like "图1 标题" or "Figure 1. Caption" immediately
-- following a standalone image paragraph by copying it into the image title.
-- A later filter can then turn that title into the final figure caption.

local function stringify_inlines(inlines)
  return pandoc.utils.stringify(inlines or {}):gsub("^%s+", ""):gsub("%s+$", "")
end

local function is_image_caption_text(text)
  if text == "" then
    return false
  end

  return text:match("^图%s*[%d一二三四五六七八九十百千]+[%s:：.、%-_].*")
      or text:match("^图%s*[%d一二三四五六七八九十百千]+$")
      or text:match("^Figure%s*[%dIVXLCivxlc]+[%s:：.、%-_].*")
      or text:match("^Figure%s*[%dIVXLCivxlc]+$")
end

local function extract_single_image(block)
  if not block or block.t ~= "Para" or #block.content ~= 1 then
    return nil
  end

  local inline = block.content[1]
  if inline.t == "Image" then
    return inline
  end

  return nil
end

function Pandoc(doc)
  local blocks = doc.blocks
  local out = {}
  local i = 1

  while i <= #blocks do
    local current = blocks[i]
    local next_block = blocks[i + 1]
    local image = extract_single_image(current)

    if image
      and next_block
      and next_block.t == "Para"
      and is_image_caption_text(stringify_inlines(next_block.content))
    then
      image.title = stringify_inlines(next_block.content)
      out[#out + 1] = current
      i = i + 2
    else
      out[#out + 1] = current
      i = i + 1
    end
  end

  return pandoc.Pandoc(out, doc.meta)
end

return { Pandoc = Pandoc }
