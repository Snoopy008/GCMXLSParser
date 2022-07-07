//
//  GCMSheetModel.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

@interface GCMSheetModel : NSObject

@property (nonatomic, copy) NSString *name;     // sheet的name属性值
@property (nonatomic, copy) NSString *sheetId;  // sheet的sheetId属性值
@property (nonatomic, copy) NSString *rid;      // sheet的r:id属性值

@end

NS_ASSUME_NONNULL_END
