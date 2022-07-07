//
//  GCMCellModel.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/5.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

typedef NS_ENUM(NSUInteger, GCMCellContentType) {
    GCMCellContentTypeBlank = 0,
    GCMCellContentTypeString,
    GCMCellContentTypeInteger,
    GCMCellContentTypeFloat,
    GCMCellContentTypeBool,
    GCMCellContentTypeError,
    GCMCellContentTypeUnknown,
};



@interface GCMCellModel : NSObject

/// 数据类型
@property (nonatomic, assign) GCMCellContentType contentType;

/// 行索引
@property (nonatomic, assign) NSInteger row;

/// 列索引
@property (nonatomic, assign) NSInteger column;

/// 列名称 "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, copy) NSString *columnName;

/// 文本值（取决于contentType）
@property (nonatomic, copy) NSString *text;

/// 数字值 （取决于contentType）
@property (nonatomic, copy) NSNumber *number;

/// sheet名称
@property (nonatomic, copy) NSString *sheetName;

@end

NS_ASSUME_NONNULL_END
