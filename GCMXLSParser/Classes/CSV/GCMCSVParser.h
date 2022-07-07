//
//  GCMCSVParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/5.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

@interface GCMCSVParser : NSObject

+ (NSMutableArray *)readCSVData:(NSString *)filePath;

@end

NS_ASSUME_NONNULL_END
