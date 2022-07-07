//
//  GCMXLSXParser.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/6.
//

#import <Foundation/Foundation.h>
#import "GCMCellModel.h"


NS_ASSUME_NONNULL_BEGIN

@interface GCMXLSXParser : NSObject

- (instancetype)initWithPath:(NSString *)path;


- (NSArray<GCMCellModel *> *)parse;

@end

NS_ASSUME_NONNULL_END
