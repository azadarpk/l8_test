<?php

namespace Tests\Unit;

use Tests\TestCase;

class ProductTest extends TestCase
{
    /**
     * A basic unit test example.
     *
     * @return void
     */
    // public function test_example()
    // {
    //     $this->assertTrue(true);
    // }

    public function test_create_form()
    {
        $response = $this->get('/products/create');
        $this->assertStatus(200);
    }

    public function test_product_create()
    {
        $product = Product::make([
            'title' => 'Dell Laptop',
            'description' => 'Dell Laptop description',
            'price' => 1500

        ]);

        $this->assertTrue($product->title);
    }
}
