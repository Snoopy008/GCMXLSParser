//
//  GCMXLSXSheetParser.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import "GCMXLSXSheetParser.h"
#import "GCMSheetCellModel.h"
#import "GCMSheetModel.h"
#import "GCMXLSXConfig.h"

@interface GCMXLSXSheetParser()<NSXMLParserDelegate>

{
    NSString *currentRowNum;        // 当前数据所处的行数
    NSString *currentIndexValue;    // 当前列中持有的索引
}
@property (nonatomic, strong) NSMutableDictionary *allColDict;  // 存储了多个sheet里的信息
@property (nonatomic, strong) NSMutableArray *colArr;   // 存储所有col的信息
@property (nonatomic, strong) GCMSheetCellModel *sheetCellModel;      // 每一个基点的数据

@end

@implementation GCMXLSXSheetParser

#pragma -mark 列的信息集合

- (NSMutableDictionary *)allColDict {
    
    if (!_allColDict) {
        _allColDict = [NSMutableDictionary dictionary];
    }
    return _allColDict;
}

/**
 *  开始解析
 *
 *  @return 解析出的数据
 */
- (id)startParse {
    
    for (int i = 0; i < self.sheetArr.count; i++) {
        
        _colArr = [NSMutableArray array];
        currentRowNum = 0;
        GCMSheetModel *sheet = self.sheetArr[i];
        NSString *sheetPath = [NSString stringWithFormat:@"%@/xl/worksheets/%@.xml",destinationPath,sheet.name.lowercaseString];
        
        self.filePath = sheetPath;
        self.xmlParser.delegate = self;
        [self.xmlParser parse];
        self.allColDict[sheet.name] = self.colArr;
        _colArr = nil;
    }
    
    return self.allColDict;
}

//准备节点
- (void)parser:(NSXMLParser *)parser didStartElement:(NSString *)elementName namespaceURI:(nullable NSString *)namespaceURI qualifiedName:(nullable NSString *)qName attributes:(NSDictionary<NSString *, NSString *> *)attributeDict{
    
    self.currentElement = elementName;
    if ([self.currentElement isEqualToString:@"row"]) {
        
        currentRowNum = attributeDict[@"r"];
    }else if ([self.currentElement isEqualToString:@"c"]) {
        
        self.sheetCellModel = [[GCMSheetCellModel alloc] init];
        self.sheetCellModel.row = currentRowNum;
        self.sheetCellModel.r = attributeDict[@"r"];
        self.sheetCellModel.s = attributeDict[@"s"]!=nil?attributeDict[@"s"]:@"";
        self.sheetCellModel.t = attributeDict[@"t"]!=nil?attributeDict[@"t"]:@"";
    }
}

// 3. 查找节点内容，可能会多次
- (void)parser:(NSXMLParser *)parser foundCharacters:(NSString *)string
{
   
    self.sheetCellModel.v = string;
    [self.colArr addObject:self.sheetCellModel];
}

////解析完一个节点
//- (void)parser:(NSXMLParser *)parser didEndElement:(NSString *)elementName namespaceURI:(nullable NSString *)namespaceURI qualifiedName:(nullable NSString *)qName{
//
//}

- (void)parserDidEndDocument:(NSXMLParser *)parser {
    
    self.xmlParser = nil;
    self.currentElement = @"";
    currentRowNum = @"";
}

@end
