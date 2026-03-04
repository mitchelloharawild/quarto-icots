--
-- body-format.lua
--
-- Applies the docx paragraph style "Body Text" to body paragraphs.
-- Renders bullet/ordered lists as raw OpenXML using the "BulletList" style
-- from the reference docx, with explicit w:ind to encode nesting depth.

-------------------------------------------------------------------------------
-- Helpers
-------------------------------------------------------------------------------

local function apply_body_text(para)
  return pandoc.Div(
    { para },
    pandoc.Attr("", {}, { ["custom-style"] = "Body Text" })
  )
end

local function empty_body_text()
  return pandoc.RawBlock(
    "openxml",
    '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>'
  )
end

local function has_custom_style(div)
  return div.attributes["custom-style"] ~= nil
    and div.attributes["custom-style"] ~= ""
end

local function has_bibliography(meta)
  local bib = meta.bibliography
  if bib == nil then return false end
  if type(bib) == "table" and bib.t == "MetaList" then
    return #bib > 0
  end
  return pandoc.utils.stringify(bib) ~= ""
end

-------------------------------------------------------------------------------
-- Raw OpenXML list rendering
--
-- Uses the "BulletList" style from the reference docx for bullet glyph and
-- base formatting. Overrides <w:ind> to set depth-based indentation:
--   depth 0 → before-text 0"   (w:left = 0 + HANGING)
--   depth 1 → before-text 0.5" (w:left = 720 + HANGING)
--   depth 2 → before-text 1.0" (w:left = 1440 + HANGING)
-------------------------------------------------------------------------------

local TWIPS_PER_LEVEL = 720  -- 0.5" per depth level
local HANGING         = 360  -- 0.25" hanging indent for the bullet label

local function xml_escape(s)
  s = s:gsub("&",  "&amp;")
  s = s:gsub("<",  "&lt;")
  s = s:gsub(">",  "&gt;")
  s = s:gsub('"',  "&quot;")
  s = s:gsub("'",  "&apos;")
  return s
end

local function inlines_to_xml(inlines)
  local parts = {}
  for _, inline in ipairs(inlines) do
    if inline.t == "Str" then
      parts[#parts + 1] = xml_escape(inline.text)
    elseif inline.t == "Space" or inline.t == "SoftBreak" then
      parts[#parts + 1] = " "
    elseif inline.t == "Emph" then
      parts[#parts + 1] = '</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">'
        .. inlines_to_xml(inline.content)
        .. '</w:t></w:r><w:r><w:t xml:space="preserve">'
    elseif inline.t == "Strong" then
      parts[#parts + 1] = '</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">'
        .. inlines_to_xml(inline.content)
        .. '</w:t></w:r><w:r><w:t xml:space="preserve">'
    elseif inline.t == "Code" then
      parts[#parts + 1] = xml_escape(inline.text)
    elseif inline.t == "RawInline" and inline.format == "openxml" then
      parts[#parts + 1] = inline.text
    else
      parts[#parts + 1] = xml_escape(pandoc.utils.stringify({ inline }))
    end
  end
  return table.concat(parts)
end

-- Emit one <w:p> using the BulletList style with depth-overridden indentation.
local function bullet_item_xml(depth, content_xml)
  local left = depth * TWIPS_PER_LEVEL + HANGING
  return string.format(
    '<w:p>'
    .. '<w:pPr>'
    ..   '<w:pStyle w:val="BulletList"/>'
    ..   '<w:ind w:left="%d" w:hanging="%d"/>'
    .. '</w:pPr>'
    .. '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
    .. '</w:p>',
    left, HANGING,
    content_xml
  )
end

-- Recursively convert a list node to flat XML strings (top-down).
-- Handles both BulletList and OrderedList (ordered uses the same style for now).
local function list_node_to_xml(node, depth)
  local parts = {}
  for _, item in ipairs(node.content) do
    for _, block in ipairs(item) do
      if block.t == "Para" or block.t == "Plain" then
        parts[#parts + 1] = bullet_item_xml(depth, inlines_to_xml(block.content))
      elseif block.t == "BulletList" or block.t == "OrderedList" then
        local nested = list_node_to_xml(block, depth + 1)
        for _, s in ipairs(nested) do
          parts[#parts + 1] = s
        end
      end
    end
  end
  return parts
end

local function list_to_rawblock(node)
  return pandoc.RawBlock("openxml", table.concat(list_node_to_xml(node, 0), "\n"))
end

-------------------------------------------------------------------------------
-- Top-down block processor
-------------------------------------------------------------------------------

local process_blocks  -- forward declaration

process_blocks = function(blocks, in_custom_div)
  local out = pandoc.List()

  for _, block in ipairs(blocks) do
    if block.t == "BulletList" or block.t == "OrderedList" then
      out:insert(empty_body_text())
      out:insert(list_to_rawblock(block))

    elseif block.t == "Header" then
      out:insert(empty_body_text())
      out:insert(block)

    elseif block.t == "Para" and not in_custom_div then
      out:insert(apply_body_text(block))

    elseif block.t == "Div" then
      if has_custom_style(block) then
        out:insert(block)
      else
        block.content = process_blocks(block.content, false)
        out:insert(block)
      end

    else
      out:insert(block)
    end
  end

  return out
end

-------------------------------------------------------------------------------
-- Filter
-------------------------------------------------------------------------------

return {
  {
    Pandoc = function(doc)
      doc.blocks = process_blocks(doc.blocks, false)
      return doc
    end,
  },

  {
    Pandoc = function(doc)
      if not has_bibliography(doc.meta) then
        return nil
      end
      doc.blocks:insert(pandoc.Header(2, { pandoc.Str("References") }))
      return doc
    end,
  },
}
