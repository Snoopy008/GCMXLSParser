//
//  GCMXLSParser.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/6.
//

#import "GCMXLSParser.h"
#import "xls.h"

static const NSInteger CELL_ID_BLANK = 0x201;

@interface GCMXLSParser()

@property (nonatomic, assign) xlsWorkBook *workBook;
@property (nonatomic, assign) NSInteger numSheets;
@property (nonatomic, assign) NSInteger activeWorkSheetID;        // keep last one active
@property (nonatomic, assign) xlsWorkSheet *activeWorkSheet;        // keep last one active
@property (nonatomic, assign) xlsSummaryInfo *summary;
@property (nonatomic, assign) BOOL iterating;
@property (nonatomic, assign) NSInteger lastRowIndex;
@property (nonatomic, assign) NSInteger lastColIndex;
@property (nonatomic, assign) NSStringEncoding encoding;

- (void)setWorkBook:(xlsWorkBook *)wb;

- (void)openSheet:(NSInteger)sheetNum;
- (void)formatContent:(GCMCellModel *)content withCell:(xlsCell *)cell;

@end

@implementation GCMXLSParser

- (instancetype)initWithPath:(NSString *)filePath
{
    self = [self init];
    if (self) {
        const char *file = [filePath cStringUsingEncoding:NSUTF8StringEncoding];
        _workBook = xls_open(file, "UTF-8");
        if(_workBook) {
            xls_parseWorkBook(_workBook);
            _numSheets = _workBook->sheets.count;
            _summary = xls_summaryInfo(_workBook);
        }
    }
    return self;
}

- (instancetype)init
{
    self = [super init];
    if (self)
    {
        _activeWorkSheetID = NSNotFound;
        _encoding = NSUTF8StringEncoding;
    }
    return self;
}

- (void)dealloc
{
    xls_close_summaryInfo(_summary);
    xls_close_WS(_activeWorkSheet);
    xls_close_WB(_workBook);
}

+ (NSString *)libaryVersion
{
    return [NSString stringWithCString:xls_getVersion() encoding:NSASCIIStringEncoding];
}

// Sheet Information
- (NSInteger)numberOfSheets
{
    return _numSheets;
}

- (NSString *)sheetNameAtIndex:(NSInteger)idx
{
    return idx < _numSheets ? [NSString stringWithCString:(char *)_workBook->sheets.sheet[idx].name encoding:_encoding] : nil;
}

- (BOOL)isSheetVisibleAtIndex:(NSUInteger)idx
{
    return idx < _numSheets ? (BOOL)_workBook->sheets.sheet[idx].visibility : NO;
}

- (void)openSheet:(NSInteger)sheetNum
{
    if(sheetNum >= _numSheets) {
        _iterating = true;
        _lastColIndex = UINT32_MAX;
        _lastRowIndex = UINT32_MAX;
    } else {
        if(sheetNum != _activeWorkSheetID) {
            _activeWorkSheetID = sheetNum;
            xls_close_WS(_activeWorkSheet);
            _activeWorkSheet = xls_getWorkSheet(_workBook, (int)sheetNum);
            xls_parseWorkSheet(_activeWorkSheet);
        }
    }
}

- (NSInteger)numberOfRowsInSheet:(NSInteger)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastrow + 1;
}

- (NSInteger)numberOfColsInSheet:(NSInteger)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastcol + 1;
}

- (GCMCellModel *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row column:(NSInteger)column
{
    GCMCellModel *content = [[GCMCellModel alloc] init];
    
    assert(row && column);
    
    [self startIteratorSheetAtIndex:NSNotFound];
    [self openSheet:sheetNum];
    
    row--;
    column--;
    
    NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
    NSUInteger numCols = _activeWorkSheet->rows.lastcol + 1;
    
    for (NSUInteger t = 0; t < numRows; t++)
    {
        xlsRow *rowP = &_activeWorkSheet->rows.row[t];
        for (NSUInteger tt=0; tt<numCols; tt++)
        {
            xlsCell    *cell = &rowP->cells.cell[tt];
            // NSLog(@"Looking for %d:%d:%d - testing %d:%d Type: 0x%4.4x  [t=%d tt=%d]", sheetNum, row, col, cell->row, cell->col, cell->id, t, tt);
            if(cell->row < row)
            {
                break;
            }
            
            if(cell->row > row)
            {
                return content;
            }
            
            if(cell->id == CELL_ID_BLANK)
            {
                // "Blank" filler cell created by libxls
                continue;
            }
            
            if(cell->col == column) {
                [self formatContent:content withCell:cell];
                return content;
            }
        }
    }
    
    return content;
}

- (GCMCellModel *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row columnName:(NSString *)columnName
{
    const char *colStr = [columnName cStringUsingEncoding:NSUTF8StringEncoding];
    if(strlen(colStr) > 2 || strlen(colStr) == 0)
    {
        return [[GCMCellModel alloc] init];
    }
    
    NSInteger col = colStr[0] - 'A';
    
    if(col < 0 || col >= 26)
    {
        return [[GCMCellModel alloc] init];
    }
    
    char c = colStr[1];
    if(c)
    {
        col *= 26;
        NSInteger col2 = c - 'A';
        if(col2 < 0 || col2 >= 26)
        {
            return [[GCMCellModel alloc] init];
        }
        col += col2;
    }
    col += 1;
    
    return [self cellInWorkSheetIndex:sheetNum row:row column:col];
}

// Iterate through all cells
- (void)startIteratorSheetAtIndex:(NSInteger)sheetIndex
{
    if(sheetIndex != NSNotFound) {
        [self openSheet:sheetIndex];
        _iterating = true;
        _lastColIndex = 0;
        _lastRowIndex = 0;
    } else {
        _iterating = false;
    }
}

- (GCMCellModel *)nextCell
{
    if(!_iterating) {
        return nil;
    }
    
    GCMCellModel *content = [[GCMCellModel alloc] init];
    
    NSUInteger rowCount = _activeWorkSheet->rows.lastrow + 1;
    NSUInteger columnCount = _activeWorkSheet->rows.lastcol + 1;
    
    if(_lastRowIndex >= rowCount) {
        return content;
    }
    
    for (NSUInteger t = _lastRowIndex; t < rowCount; t++)
    {
        xlsRow *rowP = &_activeWorkSheet->rows.row[t];
        for (NSInteger column = _lastColIndex; column < columnCount; column++)
        {
            xlsCell *cell = &rowP->cells.cell[column];
            
            if(cell->id == CELL_ID_BLANK) {
                continue;
            }
            _lastColIndex = column + 1;
            [self formatContent:content withCell:cell];
            return content;
        }
        _lastRowIndex++;
        _lastColIndex = 0;
    }
    // don't make iterator false - user can keep asking for cells, they all just be blank ones though
    return content;
}

- (void)formatContent:(GCMCellModel *)content withCell:(xlsCell *)cell
{
    NSUInteger col = cell->col;
    
    content.row = cell->row + 1;
    
    {
        content.column = col + 1;
        char colStr[3];
        if(col < 26) {
            colStr[0] = 'A' + (char)col;
            colStr[1] = '\0';
        } else {
            colStr[0] = 'A' + (char)(col/26);
            colStr[1] = 'A' + (char)(col%26);
        }
        colStr[2] = '\0';
        content.columnName = [NSString stringWithFormat:@"%s", colStr];
    }
    
    switch(cell->id) {
        case 0x0006:    //FORMULA
            // test for formula, if
            if(cell->l == 0) {
                content.contentType = GCMCellContentTypeFloat;
                content.number = [NSNumber numberWithDouble:cell->d];
            } else {
                if(!strcmp((char *)cell->str, "bool")) {
                    BOOL b = (BOOL)cell->d;
                    content.contentType = GCMCellContentTypeBool;
                    content.number = [NSNumber numberWithBool:b];
                    content.text = b ? @"YES" : @"NO";
                } else
                    if(!strcmp((char *)cell->str, "error")) {
                        // FIXME: Why do we convert the double cell->d to NSInteger?
                        NSInteger err = (NSInteger)cell->d;
                        content.contentType = GCMCellContentTypeError;
                        content.number = [NSNumber numberWithInteger:err];
                        content.text = [NSString stringWithFormat:@"%ld", (long)err];
                    } else {
                        content.contentType = GCMCellContentTypeString;
                    }
            }
            break;
        case 0x00FD:    //LABELSST
        case 0x0204:    //LABEL
            content.contentType = GCMCellContentTypeString;
            content.number = [NSNumber numberWithLong:cell->l];    // possible numeric conversion done for you
            break;
        case 0x0203:    //NUMBER
        case 0x027E:    //RK
            content.contentType = GCMCellContentTypeFloat;
            content.number = [NSNumber numberWithDouble:cell->d];
            break;
        default:
            content.contentType = GCMCellContentTypeUnknown;
            break;
    }
    
    if(!content.text) {
        content.text = [NSString stringWithCString:(char *)cell->str encoding:NSUTF8StringEncoding];
    }
}

// Summary Information
- (NSString *)appName {
    return [self stringWithChar:(char *)_summary->appName];
}

- (NSString *)author {
    return [self stringWithChar:(char *)_summary->author];
}

- (NSString *)category {
    return [self stringWithChar:(char *)_summary->category];
}

- (NSString *)comment {
    return [self stringWithChar:(char *)_summary->comment];
}

- (NSString *)company {
    return [self stringWithChar:(char *)_summary->company];
}

- (NSString *)keywords {
    return [self stringWithChar:(char *)_summary->keywords];
}

- (NSString *)lastAuthor {
    return [self stringWithChar:(char *)_summary->lastAuthor];
}

- (NSString *)manager {
    return [self stringWithChar:(char *)_summary->manager];
}

- (NSString *)subject {
    return [self stringWithChar:(char *)_summary->subject];
}

- (NSString *)title {
    return [self stringWithChar:(char *)_summary->title];
}

- (NSString *)stringWithChar:(char *)c
{
    if (!c)
    {
        return @"";
    }
    return [NSString stringWithCString:c encoding:NSUTF8StringEncoding];
}

@end
