function! s:loadXml(f, t)
  let xml = system(printf("unzip -p -- %s %s", shellescape(a:f), shellescape(a:t)))
  return webapi#xml#parse(xml)
endfunction

function! s:loadWorkbook(f)
  let relNames = {}
  let doc = s:loadXml(a:f, "xl/workbook.xml")
  let sheets = doc.childNode("sheets")
  for sheet in sheets.childNodes("sheet")
    let relNames[sheet.attr["r:id"]] = sheet.attr["name"]
  endfor
  return relNames
endfunction

function! s:loadRels(f)
  let sheetRels = {}
  let doc = s:loadXml(a:f, "xl/_rels/workbook.xml.rels")
  for rel in doc.childNodes("Relationship")
    let ms = matchlist(rel.attr["Target"], '\vworksheets/sheet(\d+).xml')
    if empty(ms)
      continue
    endif
    let sheetRels[ms[1]] = rel.attr["Id"]
  endfor
  return sheetRels
endfunction

function! s:loadSheetNames(f)
  let sheetNames = {}
  try
    let relNames = s:loadWorkbook(a:f)
    let sheetRels = s:loadRels(a:f)
    for [sheet, rel] in items(sheetRels)
      let sheetNames[sheet] = relNames[rel]
    endfor
  catch
  endtry
  return sheetNames
endfunction

function! s:loadSharedStrings(f)
  let ss = []
  try
    let doc = s:loadXml(a:f, "xl/sharedStrings.xml")
    for si in doc.childNodes("si")
      let t = si.childNode("t")
      if !empty(t)
        call add(ss, t.value())
      else
        let ts = si.findAll("t")
        call add(ss, join(map(ts, "v:val.value()"), ""))
      endif
    endfor
  catch
  endtry
  return ss
endfunction

function! s:loadSheetData(f, s)
  let ss = s:loadSharedStrings(a:f)
  let doc = s:loadXml(a:f, "xl/worksheets/sheet" . a:s . ".xml")
  let rows = doc.childNode("sheetData").childNodes("row")
  let cells = map(range(1, 256), 'map(range(1,256), "''''")')
  let aa = char2nr('A')
  for row in rows
    for col in row.childNodes("c")
      let r = col.attr["r"]
      let nv = col.childNode("v")
      let v = empty(nv) ? "" : nv.value()
      if has_key(col.attr, "t") && col.attr["t"] == "s"
        let v = ss[v]
      elseif has_key(col.attr, "s") && col.attr["s"] == "2"
        let v = strftime("%Y/%m/%d %H:%M:%S", (v - 25569) * 86400 - 32400)
      endif
      let x = char2nr(r[0]) - aa
      let y = matchstr(r, '\d\+')
      let cells[y][x+1] = v
    endfor
  endfor
  for y in range(len(cells)-1)
    let cells[y+1][0] = y + 1
  endfor
  for x in range(len(cells[0])-1)
    let nx = x / 26
    if nx == 0
      let cells[0][x+1] = nr2char(aa+x)
    else
      let cells[0][x+1] = nr2char(aa+nx-1) . nr2char(aa+x%26)
    endif
  endfor
  return cells
endfunction

function! s:fillColumns(rows)
  let rows = a:rows
  if type(rows) != 3 || type(rows[0]) != 3
    return [[]]
  endif
  let cols = len(rows[0])
  for c in range(cols)
    let m = 0
    let w = range(len(rows))
    for r in range(len(w))
      if type(rows[r][c]) == 2
        let s = string(rows[r][c])
      endif
      let w[r] = strdisplaywidth(rows[r][c])
      let m = max([m, w[r]])
    endfor
    for r in range(len(w))
      let rows[r][c] = ' ' . rows[r][c] . repeat(' ', m - w[r]) . ' '
    endfor
  endfor
  return rows
endfunction

function! s:renderSheet(f, s, sname)
  try
    let data = s:loadSheetData(a:f, a:s)
  catch
    let e = v:exception
    echohl Error | echon printf("Error while loading sheet%d: %s", a:s, e) | echohl None
    return
  endtry
  silent! execute 'tabnew' a:sname
  setlocal noswapfile buftype=nofile bufhidden=delete nowrap norightleft modifiable nolist nonumber
  let data = s:fillColumns(data) 
  let sep = "+" . join(map(copy(data[0]), 'repeat("-", len(v:val))'), '+') . "+"
  call setline(1, sep)
  let r = 2
  for row in data
    let line = join(row, '|')
    call setline(r, '|'.join(row, '|').'|')
    call setline(r + 1, sep)
    let r += 2
  endfor
  setlocal nomodifiable
endfunction

function! excelview#view(...) abort
  if a:0 > 2
    echohl Error | echon "Usage: :ExcelView [filename] {[sheet-number]}" | echohl None
    return
  endif
  let f = a:1
  let sheetNames = s:loadSheetNames(f)
  if a:0 == 1
    for s in range(1, len(sheetNames))
      call s:renderSheet(f, s, sheetNames[s])
    endfor
  else
    let s = a:2
    call s:renderSheet(f, s, sheetNames[s])
  endif
endfunction
