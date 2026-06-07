import zipfile, re, os, shutil

def save_clean(wb, out_path):
    """openpyxl save, then strip empty <v> after formulas, force recalc, drop calcChain."""
    tmp = out_path + '.tmp.xlsx'
    wb.save(tmp)
    zin = zipfile.ZipFile(tmp, 'r')
    names = zin.namelist()
    data = {n: zin.read(n) for n in names}
    zin.close(); os.remove(tmp)

    # strip empty value tags that follow a formula element (these are what Excel rejects)
    stripped = 0
    for n in list(data):
        if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            s = data[n].decode('utf-8')
            before = s.count('<v></v>') + s.count('<v/>')
            s = re.sub(r'(</f>)<v\s*/>', r'\1', s)
            s = re.sub(r'(</f>)<v\s*></v>', r'\1', s)
            s = re.sub(r'(<f[^>]*/>)<v\s*/>', r'\1', s)
            s = re.sub(r'(<f[^>]*/>)<v\s*></v>', r'\1', s)
            after = s.count('<v></v>') + s.count('<v/>')
            stripped += (before - after)
            data[n] = s.encode('utf-8')

    # force full recalc on open
    wbx = data['xl/workbook.xml'].decode('utf-8')
    if 'fullCalcOnLoad' not in wbx:
        if '<calcPr' in wbx:
            wbx = re.sub(r'<calcPr([^/]*)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
        else:
            wbx = wbx.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    data['xl/workbook.xml'] = wbx.encode('utf-8')

    # drop calcChain (Excel rebuilds it) + registrations
    if 'xl/calcChain.xml' in data:
        del data['xl/calcChain.xml']
        names = [x for x in names if x != 'xl/calcChain.xml']
        ct = data['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)
        data['[Content_Types].xml'] = ct.encode('utf-8')
        rl = data['xl/_rels/workbook.xml.rels'].decode('utf-8')
        rl = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', rl)
        data['xl/_rels/workbook.xml.rels'] = rl.encode('utf-8')

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    return stripped
