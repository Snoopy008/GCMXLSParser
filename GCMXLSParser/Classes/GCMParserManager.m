//
//  GCMParserManager.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/5.
//

#import "GCMParserManager.h"
#import "GCMCellModel.h"
#import "GCMCSVParser.h"
#import "GCMXLSParser.h"
#import "GCMXLSXParser.h"

@implementation GCMParserManager

+ (instancetype)shareInstance
{
    static id instance = nil;
    static dispatch_once_t onceToken;
    dispatch_once(&onceToken, ^{
        instance = [[self alloc] init];
    });
    return instance;
}

-(NSArray<GCMCellModel *> *)parserExcelWithPath:(NSString *)filePath
{
    
    NSString *mimeType = [[[filePath componentsSeparatedByString:@"."] lastObject] lowercaseString];
    
    if([mimeType isEqualToString:@"csv"]) {
        return [self parserExcel_CSV_WithPath:filePath];
    } else if ([mimeType isEqualToString:@"xls"]) {
        return [self parserExcel_XLS_WithPath:filePath];
    } else if ([mimeType isEqualToString:@"xlsx"]) {
        return [self parserExcel_XLSX_WithPath:filePath];
    } else {
        NSLog(@"excel格式不支持");
        return nil;
    }
}


- (NSArray<GCMCellModel *> *)parserExcel_XLS_WithPath:(NSString *)filePath
{
    NSMutableArray *cellArray = [NSMutableArray array];
    GCMXLSParser *parser = [[GCMXLSParser alloc] initWithPath:filePath];
    
    NSString *text = @"";
    
    text = [text stringByAppendingFormat:@"AppName: %@\n", parser.appName];
    text = [text stringByAppendingFormat:@"Author: %@\n", parser.author];
    text = [text stringByAppendingFormat:@"Category: %@\n", parser.category];
    text = [text stringByAppendingFormat:@"Comment: %@\n", parser.comment];
    text = [text stringByAppendingFormat:@"Company: %@\n", parser.company];
    text = [text stringByAppendingFormat:@"Keywords: %@\n", parser.keywords];
    text = [text stringByAppendingFormat:@"LastAuthor: %@\n", parser.lastAuthor];
    text = [text stringByAppendingFormat:@"Manager: %@\n", parser.manager];
    text = [text stringByAppendingFormat:@"Subject: %@\n", parser.subject];
    text = [text stringByAppendingFormat:@"Title: %@\n", parser.title];
    
    text = [text stringByAppendingFormat:@"\n\nNumber of Sheets: %ld\n", (long)parser.numberOfSheets];
    NSLog(@"%@", text);
    [parser startIteratorSheetAtIndex:0];
    
    while(YES) {
        GCMCellModel *cell = [parser nextCell];
        if(cell.contentType == GCMCellContentTypeBlank)
        {
            break;
            
        }else{
            [cellArray addObject:cell];
        }
        NSLog(@"%@", cell);
    }
    return [NSArray arrayWithArray:cellArray];
}

- (NSArray<GCMCellModel *> *)parserExcel_XLSX_WithPath:(NSString *)filePath
{
    GCMXLSXParser *parser = [[GCMXLSXParser alloc] initWithPath:filePath];
    
    NSArray *cellArray = [parser parse];
    NSLog(@"%@",cellArray);
    
    return cellArray;

}

- (NSArray<GCMCellModel *> *)parserExcel_CSV_WithPath:(NSString *)filePath
{
    NSMutableArray *result = [GCMCSVParser readCSVData:filePath];
    NSMutableArray<GCMCellModel *> *resultArray = [NSMutableArray array];
    int row = 0;
    for(NSArray *item in result) {
        
        for(int i=0;i<item.count; i++) {
            GCMCellModel *cell = [[GCMCellModel alloc] init];
            cell.row = row;
            cell.text = item[i];
            cell.column = i;
            cell.contentType = GCMCellContentTypeString;
            [resultArray addObject:cell];
            NSLog(@"%@",cell) ;
        }
        row++;
        
    }
    return [NSArray arrayWithArray:resultArray];
    
}




@end
