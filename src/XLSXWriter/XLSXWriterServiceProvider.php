<?php

namespace XLSXWriter;

use Illuminate\Support\ServiceProvider;

class XLSXWriterServiceProvider extends ServiceProvider
{
    /**
     * Register the service provider.
     *
     * @return void
     */
    public function register()
    {
        $this->app->singleton('XLSXWriter', XLSXWriter::class);
    }
}
