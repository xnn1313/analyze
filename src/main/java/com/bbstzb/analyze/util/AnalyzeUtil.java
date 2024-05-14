package com.bbstzb.analyze.util;
import com.atilika.kuromoji.ipadic.Token;
import com.atilika.kuromoji.ipadic.Tokenizer;

import java.util.List;

public class AnalyzeUtil {
    // 计算句子中的动词数量
    public static int countVerbs(String sentence) {
        Tokenizer tokenizer = new Tokenizer(); // 初始化分词器
        List<Token> tokens = tokenizer.tokenize(sentence);

        int verbCount = 0;
        boolean lastVerbIsTeForm = false;
        for (Token token : tokens) {
            String partOfSpeechLevel1 = token.getPartOfSpeechLevel1(); // 获取词性层级1
            String partOfSpeechLevel2 = token.getPartOfSpeechLevel2(); // 获取词性层级2

            if ("動詞".equals(partOfSpeechLevel1)) {
                // 检查动词的细节层级，如是否为自立或非自立
                if ("自立".equals(partOfSpeechLevel2)) {
                    if (token.getSurface().endsWith("て") || token.getSurface().endsWith("で")) {
                        lastVerbIsTeForm = true; // 检查是否为て形
                    } else {
                        if (!lastVerbIsTeForm) {
                            verbCount++; // 计数非て形的自立动词
                        }
                        lastVerbIsTeForm = false;
                    }
                } else if ("非自立".equals(partOfSpeechLevel2)) {
                    if (!lastVerbIsTeForm) {
                        // 非自立动词通常不会单独使用，因此在这里不计数
                    }
                    lastVerbIsTeForm = false; // 重置状态
                }
            }
        }
        return verbCount;
    }

    // 判断句子是否为复句
    public static boolean isComplexSentence(String sentence) {
        return countVerbs(sentence) >= 2; // 判断动词数量是否大于等于2
    }
}
