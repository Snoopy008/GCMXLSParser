//
//  GCMXLSXSheetParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import "GCMXLSXBaseParser.h"
#import "GCMXLSXParserProtocol.h"

NS_ASSUME_NONNULL_BEGIN

@interface GCMXLSXSheetParser : GCMXLSXBaseParser<GCMXLSXParserProtocol>

@property (nonatomic, strong) NSArray *sheetArr;    // 从workbook解析出的sheet的名称等信息

@end

NS_ASSUME_NONNULL_END
