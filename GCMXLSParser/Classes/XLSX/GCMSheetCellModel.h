//
//  GCMSheetCellModel.h
//  GCMXLSParser
//
//  Created by samuel on 2022/7/7.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

@interface GCMSheetCellModel : NSObject

@property (nonatomic, copy) NSString *r;    // sheet中的c节点中属性r
@property (nonatomic, copy) NSString *s;    // sheet中的c节点中属性s
@property (nonatomic, copy) NSString *t;    // sheet中的c节点中属性t  数值类型
@property (nonatomic, copy) NSString *v;    // sheet中的c节点中属性v
@property (nonatomic, copy) NSString *row;  // sheet中的c列所处的行号

@end

NS_ASSUME_NONNULL_END
