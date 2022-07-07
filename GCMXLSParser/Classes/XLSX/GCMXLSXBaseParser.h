//
//  GCMXLSXBaseParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

@interface GCMXLSXBaseParser : NSObject

@property (nonatomic, copy) NSString *filePath;    // 要解析的文件路径
@property (nonatomic, strong) NSXMLParser *xmlParser;   // xml解析器
@property (nonatomic, copy) NSString *currentElement;   // 当前节点的名称

@end

NS_ASSUME_NONNULL_END
