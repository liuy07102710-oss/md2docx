-- Promote a paragraph like "图1 标题" or "Figure 1. Caption" immediately
-- following a standalone image paragraph by copying it into the image title.
-- A later filter can then turn that title into the final figure caption.

local function stringify_inlines(inlines)
  local text = pandoc.utils.stringify(inlines or {})
  text = text:gsub("^%s+", "")
  text = text:gsub("%s+$", "")
  return text
end

local CAPTION_LINE_SEPARATOR = "<<<MD2DOCX_IMAGE_CAPTION_BREAK>>>"

local function is_image_caption_text(text)
  if text == "" then
    return false
  end

  return text:match("^图%s*[%d一二三四五六七八九十百千]+[%s:：.、%-_].*")
      or text:match("^图%s*[%d一二三四五六七八九十百千]+$")
      or text:match("^Figure%s*[%dIVXLCivxlc]+[%s:：.、%-_].*")
      or text:match("^Figure%s*[%dIVXLCivxlc]+$")
      or text:match("^Fig%.?%s*[%dIVXLCivxlc]+[%s:：.、%-_].*")
      or text:match("^Fig%.?%s*[%dIVXLCivxlc]+$")
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
      local caption_lines = { stringify_inlines(next_block.content) }
      local j = i + 2

      while j <= #blocks
        and blocks[j].t == "Para"
        and is_image_caption_text(stringify_inlines(blocks[j].content))
      do
        caption_lines[#caption_lines + 1] = stringify_inlines(blocks[j].content)
        j = j + 1
      end

      image.attributes["md2docx-caption"] = table.concat(caption_lines, CAPTION_LINE_SEPARATOR)
      out[#out + 1] = current
      i = j
    else
      out[#out + 1] = current
      i = i + 1
    end
  end

  return pandoc.Pandoc(out, doc.meta)
end

return { Pandoc = Pandoc }
