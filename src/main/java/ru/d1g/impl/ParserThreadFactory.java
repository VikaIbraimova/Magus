package ru.d1g.impl;

import org.springframework.beans.factory.FactoryBean;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Component;
import ru.d1g.Utils;

/**
 * Created by A on 12.05.2017.
 */
@Component
public class ParserThreadFactory implements FactoryBean<ParserThread> {

    private final Utils utils;
    private final Parser parser;

    @Autowired
    public ParserThreadFactory(Utils utils,@Lazy Parser parser) {
        this.utils = utils;
        this.parser = parser;
    }

    @Override
    public ParserThread getObject() throws Exception {
        return new ParserThread(utils, parser);
    }

    @Override
    public Class<?> getObjectType() {
        return ParserThread.class;
    }

    @Override
    public boolean isSingleton() {
        return false;
    }
}
