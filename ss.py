import csv


class SSRow:
  def __init__(self, cols, values = None):
    self.cols = cols
    self.data = [None] * cols
    self.colidx = 0
    if values is not None:
      for i in range(min(len(values), self.cols)):
        self.data[i] = values[i]

  def __getitem__(self, i):
    return self.data[i]

  def __setitem__(self, i, v):
    self.data[i] = v

  def __len__(self):
    return self.cols

  def __iter__(self):
    return self

  def next(self):
    if self.colidx >= self.cols:
      self.colidx = 0
      raise StopIteration
    colidx = self.colidx
    self.colidx += 1
    return self.data[colidx]



class SSArea:
  def __init__(self, cols, rows = 0):
    self.cols = cols
    self.rows = 0
    self.data = []
    self.fieldsep = ','
    self.linesep = "\n"
    self.rowidx = 0
    if rows > 0:
      for row in range(rows):
        self.addRow()

  def addRow(self, values = None):
    self.data.append(SSRow(self.cols, values))
    self.rows += 1

  def __getitem__(self, i):
    return self.data[i]

  def __iter__(self):
    return self

  def next(self):
    if self.rowidx >= self.rows:
      self.rowidx = 0
      raise StopIteration
    rowidx = self.rowidx
    self.rowidx += 1
    return self.data[rowidx]

  def toVector(self):
    r = []
    for row in self.data:
      newrow = []
      for cell in row:
        newrow.append(cell)
      r.append(newrow)
    return r

  def frame(self, rowstart, colstart, rows, cols):
    r = SSArea(cols)
    for i in range(rows):
      r.addRow()
      for j in range(cols):
        r[i][j] = self[i+rowstart][j+colstart]
    return r
    

  def loadCsv(self, fn):
    self.rows = 0
    self.cols = 0
    self.data = []
    with open(fn) as f:
      r = csv.reader(f, delimiter=self.fieldsep)
      for row in r:
        i = self.rows
        if i == 0:
          self.cols = len(row)
        self.addRow()
        j = 0
        for v in row:
          self[i][j] = v
          j += 1

  def writeCsv(self, fn):
    with open(fn, 'wb') as f:
      w = csv.writer(f, delimiter=self.fieldsep, quotechar='"', quoting=csv.QUOTE_MINIMAL)
      for row in self.data:
        w.writerow(row)

  def __str__(self):
    s = ''
    for i in range(self.rows):
      if i > 0:
        s += self.linesep
      r = self[i]
      for j in range(self.cols):
        if j == 0:
          s += str(r[j])
        else:
          s += self.fieldsep + str(r[j])
    return s
          
