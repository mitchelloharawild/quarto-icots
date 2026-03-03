--
-- Output order: Title (Heading 1), Authors (Normal, centred), Abstract (Abstract)
--
-- Author line format (centred):
--   __Author1__^1^ and Author2^2^
--   ^1^Affiliation 1
--   ^2^Affiliation 2
--   email@address
--
-- All other metadata fields (date, keywords, etc.) are suppressed.

-------------------------------------------------------------------------------
-- OpenXML helpers
-------------------------------------------------------------------------------

--- Escape a string for inclusion in XML character data.
local function xml_escape(s)
  s = s:gsub("&",  "&amp;")
  s = s:gsub("<",  "&lt;")
  s = s:gsub(">",  "&gt;")
  s = s:gsub("\"", "&quot;")
  return s
end

--- Convert a list of pandoc Inline elements to OpenXML run strings.
--- Handles: Str, Space, SoftBreak, LineBreak, Strong, Emph,
---          Underline, Strikeout, Superscript, Subscript, Span.
local function inlines_to_xml(inlines, rpr_extra)
  rpr_extra = rpr_extra or ""
  local parts = {}

  local function run(text, rpr)
    -- Wrap text in a <w:r> with optional run properties.
    rpr = rpr or ""
    if rpr ~= "" then
      table.insert(parts, "<w:r><w:rPr>" .. rpr .. "</w:rPr><w:t xml:space=\"preserve\">"
        .. xml_escape(text) .. "</w:t></w:r>")
    else
      table.insert(parts, "<w:r><w:t xml:space=\"preserve\">"
        .. xml_escape(text) .. "</w:t></w:r>")
    end
  end

  -- Forward declaration so nested calls work.
  local process

  process = function(ils, rpr)
    for _, il in ipairs(ils) do
      local t = il.t
      if t == "Str" then
        run(il.text, rpr)
      elseif t == "Space" or t == "SoftBreak" then
        run(" ", rpr)
      elseif t == "LineBreak" then
        table.insert(parts, "<w:r><w:br/></w:r>")
      elseif t == "Strong" then
        process(il.content, rpr .. "<w:b/><w:bCs/>")
      elseif t == "Emph" then
        process(il.content, rpr .. "<w:i/><w:iCs/>")
      elseif t == "Underline" then
        process(il.content, rpr .. "<w:u w:val=\"single\"/>")
      elseif t == "Strikeout" then
        process(il.content, rpr .. "<w:strike/>")
      elseif t == "Superscript" then
        process(il.content, rpr .. "<w:vertAlign w:val=\"superscript\"/>")
      elseif t == "Subscript" then
        process(il.content, rpr .. "<w:vertAlign w:val=\"subscript\"/>")
      elseif t == "Span" then
        -- Check for custom-style (character style) or underline class
        local cs = il.attributes["custom-style"]
        if cs == "Underline" then
          process(il.content, rpr .. "<w:u w:val=\"single\"/>")
        elseif cs then
          process(il.content, rpr .. "<w:rStyle w:val=\"" .. xml_escape(cs) .. "\"/>")
        else
          process(il.content, rpr)
        end
      else
        -- Fallback: stringify unknown inline types
        run(pandoc.utils.stringify({ il }), rpr)
      end
    end
  end

  process(inlines, rpr_extra)
  return table.concat(parts)
end

--- Build a centred OpenXML paragraph with NO named paragraph style.
local function centred_para(inlines)
  local pPr = "<w:pPr><w:jc w:val=\"center\"/></w:pPr>"
  local runs = inlines_to_xml(inlines)
  local xml  = "<w:p>" .. pPr .. runs .. "</w:p>"
  return pandoc.RawBlock("openxml", xml)
end

--- Build a centred OpenXML paragraph with a named paragraph style.
local function centred_styled_para(inlines, style)
  local pPr = table.concat({
    "<w:pPr>",
      "<w:pStyle w:val=\"" .. xml_escape(style) .. "\"/>",
      "<w:jc w:val=\"center\"/>",
    "</w:pPr>",
  })
  local runs = inlines_to_xml(inlines)
  local xml  = "<w:p>" .. pPr .. runs .. "</w:p>"
  return pandoc.RawBlock("openxml", xml)
end

--- Build a left-aligned OpenXML paragraph with a named paragraph style (no centring).
local function styled_raw_para(inlines, style)
  local pPr = "<w:pPr><w:pStyle w:val=\"" .. xml_escape(style) .. "\"/></w:pPr>"
  local runs = inlines_to_xml(inlines)
  local xml  = "<w:p>" .. pPr .. runs .. "</w:p>"
  return pandoc.RawBlock("openxml", xml)
end

-------------------------------------------------------------------------------
-- Pandoc AST helpers
-------------------------------------------------------------------------------

--- Wrap inlines in a Div that carries a custom docx paragraph style.
local function styled_para(inlines, style)
  local para = pandoc.Para(inlines)
  return pandoc.Div(
    { para },
    pandoc.Attr("", {}, { ["custom-style"] = style })
  )
end

--- Stringify a MetaInlines / MetaBlocks / string value safely.
local function meta_str(val)
  if val == nil then return "" end
  if type(val) == "string" then return val end
  return pandoc.utils.stringify(val)
end

--- Convert a MetaInlines value to a list of Inline elements.
local function meta_inlines(val)
  if val == nil then return {} end
  if type(val) == "table" and val.t == "MetaInlines" then
    return val
  end
  return { pandoc.Str(pandoc.utils.stringify(val)) }
end

-------------------------------------------------------------------------------
-- Build the author block
-------------------------------------------------------------------------------
local function build_author_paragraphs(authors)
  local parsed      = {}
  local affil_index = {}
  local affil_list  = {}
  local next_idx    = 1

  for _, author in ipairs(authors) do
    -- by-author entries are MetaMaps with Quarto's normalized schema:
    --   .name.literal  (MetaInlines)
    --   .email         (MetaInlines, optional)
    --   .affiliations  (MetaList of MetaMaps with .name)
    --   .attributes.corresponding (MetaBool, optional)
    local name  = meta_str(author.name and author.name.literal or author.name or "")
    local email = author.email and meta_str(author.email) or ""

    local author_affil_idxs = {}
    if author.affiliations then
      for _, affil in ipairs(author.affiliations) do
        local aname = meta_str(affil.name or "")
        if aname ~= "" then
          if not affil_index[aname] then
            affil_index[aname] = next_idx
            affil_list[next_idx] = aname
            next_idx = next_idx + 1
          end
          table.insert(author_affil_idxs, affil_index[aname])
        end
      end
    end

    table.insert(parsed, {
      name  = name,
      email = email,
      idxs  = author_affil_idxs,
    })
  end

  -- First author with an email is the corresponding author (underlined)
  local email_author_pos = nil
  local email_str        = ""
  for i, a in ipairs(parsed) do
    if a.email ~= "" then
      email_author_pos = i
      email_str = a.email
      break
    end
  end

  -- ── Name line ─────────────────────────────────────────────────────────────
  local name_line = {}
  for i, a in ipairs(parsed) do
    if i > 1 then
      table.insert(name_line, pandoc.Space())
      table.insert(name_line, pandoc.Str("and"))
      table.insert(name_line, pandoc.Space())
    end

    local name_inlines = { pandoc.Str(a.name) }

    -- Underline corresponding author
    if i == email_author_pos then
      name_inlines = { pandoc.Underline(name_inlines) }
    end

    for _, inl in ipairs(name_inlines) do
      table.insert(name_line, inl)
    end

    -- Superscript affiliation numbers
    for _, idx in ipairs(a.idxs) do
      table.insert(name_line, pandoc.Superscript({ pandoc.Str(tostring(idx)) }))
    end
  end

  -- ── Assemble paragraph list ───────────────────────────────────────────────
  local paras = {}

  table.insert(paras, { inlines = name_line, style = "Normal" })

  for idx, aname in ipairs(affil_list) do
    local inlines = {
      pandoc.Superscript({ pandoc.Str(tostring(idx)) }),
      pandoc.Str(aname),
    }
    table.insert(paras, { inlines = inlines, style = "Normal" })
  end

  -- Email line
  if email_str ~= "" then
    table.insert(paras, { inlines = { pandoc.Str(email_str) }, style = "Normal" })
  end

  return paras
end

-------------------------------------------------------------------------------
-- Main filter
-------------------------------------------------------------------------------

local extracted = {
  title    = nil,
  authors  = nil,
  abstract = nil,
}

function Meta(meta)
  extracted.title    = meta.title
  extracted.authors  = meta["by-author"]   -- denormalized: affiliations already embedded
  extracted.abstract = meta.abstract

  -- Return empty meta to suppress all default metadata rendering
  return pandoc.Meta({})
end

function Pandoc(doc)
  local blocks = pandoc.List()

  -- Title
  if extracted.title then
    blocks:insert(styled_raw_para(meta_inlines(extracted.title), "Heading1"))
    blocks:insert(pandoc.RawBlock("openxml", "<w:p/>"))
  end

  -- Authors (centred via raw OpenXML)
  if extracted.authors and #extracted.authors > 0 then
    for _, p in ipairs(build_author_paragraphs(extracted.authors)) do
      blocks:insert(centred_para(p.inlines))
    end
    blocks:insert(centred_para({}))
  end

  -- Abstract
  if extracted.abstract then
    local abs_inlines
    if type(extracted.abstract) == "table" and
       extracted.abstract.t == "MetaBlocks" then
      -- Flatten block content to inlines
      abs_inlines = {}
      for _, blk in ipairs(extracted.abstract) do
        if blk.t == "Para" or blk.t == "Plain" then
          for _, inl in ipairs(blk.content) do
            table.insert(abs_inlines, inl)
          end
        end
      end
    else
      abs_inlines = { pandoc.Str(meta_str(extracted.abstract)) }
    end
    blocks:insert(styled_raw_para(abs_inlines, "Abstract"))
  end

  -- Original body
  for _, blk in ipairs(doc.blocks) do
    blocks:insert(blk)
  end

  return pandoc.Pandoc(blocks, doc.meta)
end
