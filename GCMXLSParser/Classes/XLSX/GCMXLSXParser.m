//
//  GCMXLSXParser.m
//  GCMXLSParser
//
//  Created by samuel on 2022/7/6.
//

#import "GCMXLSXParser.h"
#import "GCMXLSXConfig.h"
#import "SSZipArchive.h"
#import "GCMCellModel.h"
#import "GCMXLSXWorkbookParser.h"
#import "GCMXLSXSheetParser.h"
#import "GCMXLSXContentParser.h"


@interface GCMXLSXParser()

@property (nonatomic, strong) GCMXLSXWorkbookParser *wbParser;    // workbookParser的解析器
@property (nonatomic, strong) GCMXLSXSheetParser *sheetParser;    // sheetParser的解析器
@property (nonatomic, strong) GCMXLSXContentParser *contentParser;// 内容的解析器

@property (nonatomic, copy) NSString *parserPath;

@end

@implementation GCMXLSXParser


- (instancetype)initWithPath:(NSString *)path
{
    self = [super init];
    if (self) {
        self.parserPath = path;
    }
    return self;
}


- (NSArray<GCMCellModel *> *)parse {
    
    // 解压xlsx文件
    [SSZipArchive unzipFileAtPath:self.parserPath toDestination:destinationPath];
    
    // 解析workbook.xml文件
    NSArray *sheetArr = [self parseWorkbook];
    
    // 解析sheet文件，获取具体值在sharedString.xml中的索引
    NSDictionary *dict = [self parseSheetWithSheetArr:sheetArr];
    // 解析最终结果
    NSArray<GCMCellModel *> *result = [self parseSharedStringsWithAllColDict:dict];
    
    [self deleteCacheFile];
    
    return result;
}

- (void)deleteCacheFile
{
    NSFileManager *fileManager = [NSFileManager defaultManager];
    if([fileManager fileExistsAtPath:destinationPath]){
        NSError *error;
        if([fileManager removeItemAtPath:destinationPath error:&error]&& !error){
//                    NSLog(@"file remove success....");
        }else if(error){
            NSLog(@"file remove failure.... error:%@",error);
        }
    }
}

/**
 *  解析workbook.xml中的数据
 *
 *  @return sheet的集合数组
 */
- (NSArray *)parseWorkbook{

    _wbParser = [[GCMXLSXWorkbookParser alloc] init];
    return [_wbParser startParse];
}


/**
 *  解析每个sheet文件
 *
 *  @return 解析出的数据数组
 */
- (NSDictionary *)parseSheetWithSheetArr:(NSArray *)arr {
  
    _sheetParser = [[GCMXLSXSheetParser alloc] init];
    [_sheetParser setSheetArr:arr];
    return [_sheetParser startParse];
}

/**
 *  根据索引从sharedStrings.xml中解析数据
 *
 *  @param dict 从sheet中解析出来的索引数据
 *
 *  @return 对应的数据数组
 */
- (NSArray<GCMCellModel *> *)parseSharedStringsWithAllColDict:(NSDictionary *)dict {

    _contentParser = [[GCMXLSXContentParser alloc] init];
    [_contentParser setAllColDict:dict];
    return [_contentParser startParse];
}


@end
