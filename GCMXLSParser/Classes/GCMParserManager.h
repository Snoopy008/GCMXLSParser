//
//  GCMParserManager.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/5.
//

#import <Foundation/Foundation.h>
#import "GCMCellModel.h"

NS_ASSUME_NONNULL_BEGIN

//
@interface GCMParserManager : NSObject

+ (instancetype)shareInstance;


/// 根据文件后缀解析数据（支持xls、xlsx、csv后缀）
/// @param filePath 文件路径
-(NSArray<GCMCellModel *> *) parserExcelWithPath:(NSString *)filePath;


/// 解析xls文件数据
/// @param filePath 文件路径
-(NSArray<GCMCellModel *> *) parserExcel_XLS_WithPath:(NSString *)filePath;


/// 解析xlsx文件数据
/// @param filePath 文件路径
-(NSArray<GCMCellModel *> *) parserExcel_XLSX_WithPath:(NSString *)filePath;


/// 解析csv文件数据
/// @param filePath 文件路径
- (NSArray<GCMCellModel *> *)parserExcel_CSV_WithPath:(NSString *)filePath;

@end

NS_ASSUME_NONNULL_END
