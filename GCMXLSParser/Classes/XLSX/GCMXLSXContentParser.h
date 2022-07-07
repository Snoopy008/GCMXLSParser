//
//  GCMXLSXContentParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import "GCMXLSXBaseParser.h"
#import "GCMXLSXParserProtocol.h"

NS_ASSUME_NONNULL_BEGIN

@interface GCMXLSXContentParser : GCMXLSXBaseParser<GCMXLSXParserProtocol>

@property (nonatomic, strong) NSDictionary *allColDict;    // 从所有sheet解析出的col的信息

@end

NS_ASSUME_NONNULL_END
