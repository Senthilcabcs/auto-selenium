<?php

namespace Database\Seeders;

use App\Models\User;
use Illuminate\Database\Seeder;
use Illuminate\Support\Facades\Hash;

class CiUserSeeder extends Seeder
{
    public function run(): void
    {
        User::updateOrCreate(
            ['email' => env('CI_TEST_EMAIL')],
            [
                'name' => 'CI Selenium User',
                'password' => Hash::make(env('CI_TEST_PASSWORD')),
            ]
        );
    }
}